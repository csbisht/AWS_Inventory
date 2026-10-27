<#
.SYNOPSIS
    Tests whether a user can RDP into multiple Windows servers using only built-in
    Windows tools (mstsc, cmdkey, quser, logoff). No FreeRDP needed.

.DESCRIPTION
    Mode Session (default) - real RDP test:
        1. Checks TCP 3389
        2. Caches creds with cmdkey and launches mstsc
        3. Polls the target with 'quser' until the test user's session appears (= access YES)
        4. Logs the test session off and closes mstsc
       Requirements: run from a desktop session (mstsc is GUI), and the account
       RUNNING the script must be able to run 'quser'/'logoff' against the targets
       (normally local admin on the target). Don't touch the mouse/keyboard while it runs.

    Mode Precheck - headless, no login:
        Port 3389 + validates the password + checks 'Remote Desktop Users' /
        'Administrators' membership over WinRM. Direct membership only, nested
        groups are not expanded.

.EXAMPLE
    .\Test-RdpAccess.ps1 -ServerFile .\servers.txt
.EXAMPLE
    .\Test-RdpAccess.ps1 -Servers srv01,srv02 -Mode Precheck
#>
[CmdletBinding()]
param(
    [string[]]$Servers,
    [string]$ServerFile,
    [ValidateSet('Session','Precheck')]
    [string]$Mode = 'Session',
    [int]$TimeoutSec = 45,
    [int]$Port = 3389,
    [string]$OutCsv = ".\rdp-test-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
)

# ====================== HARDCODED TEST CREDENTIALS ======================
# Domain account : "DOMAIN\testuser" or "testuser@corp.example.com"
# Local account  : ".\localuser"  (auto-expanded to SERVERNAME\localuser per host)
$User = "DOMAIN\testuser"
$Pass = "P@ssw0rd!"
# ========================================================================

if ($ServerFile) { $Servers += Get-Content $ServerFile | Where-Object { $_.Trim() -and $_ -notmatch '^\s*#' } }
if (-not $Servers) { throw "No servers supplied. Use -Servers or -ServerFile." }

function Get-UserForServer([string]$server) {
    if ($User -like '.\*') { "$($server.Split('.')[0])\$($User.Substring(2))" } else { $User }
}

function Get-ShortName([string]$u) {
    if ($u -like '*\*') { $u = $u.Split('\')[1] }
    if ($u -like '*@*') { $u = $u.Split('@')[0] }
    $u.Substring(0, [Math]::Min(20, $u.Length))   # quser truncates at 20 chars
}

function Get-RdpSessionIds([string]$server, [string]$shortName) {
    $out = quser /server:$server 2>$null
    foreach ($line in $out | Select-Object -Skip 1) {
        if ($line -match '^\s*>?(?<user>\S+)\s+(?:(?<sess>\S+)\s+)?(?<id>\d+)\s+(?<state>\S+)') {
            if ($Matches.user -ieq $shortName) { [int]$Matches.id }
        }
    }
}

# Suppress the "identity of the remote computer cannot be verified" prompt during the run
$tsKey = 'HKCU:\Software\Microsoft\Terminal Server Client'
$origAuthLevel = $null
if ($Mode -eq 'Session') {
    if (-not (Test-Path $tsKey)) { New-Item $tsKey -Force | Out-Null }
    $origAuthLevel = (Get-ItemProperty $tsKey -Name AuthenticationLevelOverride -ErrorAction SilentlyContinue).AuthenticationLevelOverride
    Set-ItemProperty $tsKey -Name AuthenticationLevelOverride -Value 0 -Type DWord
}

try {
$results = foreach ($server in $Servers) {
    $server  = $server.Trim()
    $rdpUser = Get-UserForServer $server
    $short   = Get-ShortName $rdpUser
    Write-Host "`n==> $server  (as $rdpUser)" -ForegroundColor Cyan

    $r = [ordered]@{ Server = $server; User = $rdpUser; PortOpen = $false; RdpAccess = 'NO'; Reason = '' }

    # 1. Port check
    try {
        $r.PortOpen = [bool](Test-NetConnection -ComputerName $server -Port $Port -WarningAction SilentlyContinue).TcpTestSucceeded
    } catch {}
    if (-not $r.PortOpen) {
        $r.Reason = "Port $Port closed/unreachable"
        Write-Host "   RDP access: NO  $($r.Reason)" -ForegroundColor Red
        [pscustomobject]$r; continue
    }

    if ($Mode -eq 'Session') {
        # 2. Real login via mstsc, verified with quser
        $existing = @(Get-RdpSessionIds $server $short)
        $target = "TERMSRV/$server"
        cmdkey /generic:$target /user:$rdpUser /pass:$Pass | Out-Null
        $proc = Start-Process mstsc -ArgumentList "/v:$server" -PassThru

        $newId = $null
        $deadline = (Get-Date).AddSeconds($TimeoutSec)
        while ((Get-Date) -lt $deadline -and -not $newId) {
            Start-Sleep -Seconds 3
            $newId = Get-RdpSessionIds $server $short | Where-Object { $_ -notin $existing } | Select-Object -First 1
        }

        if ($newId) {
            $r.RdpAccess = 'YES'
            Start-Sleep -Seconds 2
            logoff $newId /server:$server 2>$null
        } elseif ($existing.Count) {
            $r.Reason = "User already had a session (id $($existing -join ',')) - couldn't verify a new login"
        } else {
            $r.Reason = "No session within ${TimeoutSec}s (bad creds / no RDP logon right / NLA or policy block)"
        }

        if ($proc -and -not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
        Get-Process mstsc -ErrorAction SilentlyContinue | Where-Object { $_.Id -eq $proc.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
        cmdkey /delete:$target | Out-Null
    }
    else {
        # 2. Precheck: password valid?
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $credOk = $false
        try {
            if ($rdpUser -like "$($server.Split('.')[0])\*") {
                $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Machine', $server)
            } else {
                $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain')
            }
            $credOk = $ctx.ValidateCredentials((Get-ShortName $rdpUser), $Pass)
        } catch { $r.Reason = "Cred check error: $($_.Exception.Message)" }

        if (-not $credOk) {
            if (-not $r.Reason) { $r.Reason = 'Invalid username/password' }
        } else {
            # 3. RDP group membership (runs as the account executing the script)
            try {
                $members = Invoke-Command -ComputerName $server -ErrorAction Stop -ScriptBlock {
                    @(Get-LocalGroupMember 'Remote Desktop Users' -ErrorAction SilentlyContinue) +
                    @(Get-LocalGroupMember 'Administrators'       -ErrorAction SilentlyContinue) |
                        ForEach-Object { $_.Name }
                }
                $hit = $members | Where-Object { (Get-ShortName $_) -ieq (Get-ShortName $rdpUser) }
                if ($hit) { $r.RdpAccess = 'YES' }
                else      { $r.Reason = "Creds OK, but user not directly in Remote Desktop Users/Administrators (check nested groups)" }
            } catch { $r.Reason = "Creds OK, WinRM group check failed: $($_.Exception.Message)" }
        }
    }

    $color = if ($r.RdpAccess -eq 'YES') { 'Green' } else { 'Red' }
    Write-Host "   RDP access: $($r.RdpAccess)  $($r.Reason)" -ForegroundColor $color
    [pscustomobject]$r
}
}
finally {
    if ($Mode -eq 'Session') {
        if ($null -eq $origAuthLevel) { Remove-ItemProperty $tsKey -Name AuthenticationLevelOverride -ErrorAction SilentlyContinue }
        else { Set-ItemProperty $tsKey -Name AuthenticationLevelOverride -Value $origAuthLevel -Type DWord }
    }
}

# ============================== SUMMARY ==============================
$yes = @($results | Where-Object RdpAccess -eq 'YES')
$no  = @($results | Where-Object RdpAccess -ne 'YES')

Write-Host "`n========== RDP ACCESS SUMMARY ==========" -ForegroundColor Yellow
Write-Host "HAS ACCESS ($($yes.Count)):" -ForegroundColor Green
$yes | ForEach-Object { Write-Host "  [YES] $($_.Server)" -ForegroundColor Green }
Write-Host "NO ACCESS ($($no.Count)):" -ForegroundColor Red
$no  | ForEach-Object { Write-Host "  [NO]  $($_.Server)  -> $($_.Reason)" -ForegroundColor Red }

$results | Format-Table Server, User, PortOpen, RdpAccess, Reason -AutoSize
$results | Export-Csv -Path $OutCsv -NoTypeInformation
Write-Host "Results saved to $OutCsv" -ForegroundColor Yellow
