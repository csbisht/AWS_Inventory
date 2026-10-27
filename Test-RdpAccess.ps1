<#
.SYNOPSIS
    Tests RDP login for a user across multiple Windows servers, one at a time,
    using built-in Windows tools (mstsc + cmdkey).

.DESCRIPTION
    Per server:
      1. Checks TCP 3389
      2. Launches mstsc with the hardcoded creds
      3. Detects a successful login from the mstsc window title
         ("<server> - Remote Desktop Connection")
      4. Keeps the session OPEN and waits until you LOG OFF inside the remote
         session (Start > user icon > Sign out). mstsc then closes itself and
         the script moves to the next server.
    If login doesn't happen within -TimeoutSec, mstsc is closed and the server is marked NO.

.EXAMPLE
    .\Test-RdpAccess.ps1 -ServerFile .\servers.txt
.EXAMPLE
    .\Test-RdpAccess.ps1 -Servers srv01,srv02 -TimeoutSec 60
#>
[CmdletBinding()]
param(
    [string[]]$Servers,
    [string]$ServerFile,
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

# Suppress the "identity of the remote computer cannot be verified" prompt during the run
$tsKey = 'HKCU:\Software\Microsoft\Terminal Server Client'
if (-not (Test-Path $tsKey)) { New-Item $tsKey -Force | Out-Null }
$origAuthLevel = (Get-ItemProperty $tsKey -Name AuthenticationLevelOverride -ErrorAction SilentlyContinue).AuthenticationLevelOverride
Set-ItemProperty $tsKey -Name AuthenticationLevelOverride -Value 0 -Type DWord

try {
$results = foreach ($server in $Servers) {
    $server  = $server.Trim()
    $rdpUser = Get-UserForServer $server
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

    # 2. Launch mstsc with cached creds
    $target = "TERMSRV/$server"
    cmdkey /generic:$target /user:$rdpUser /pass:$Pass | Out-Null
    $proc  = Start-Process mstsc -ArgumentList "/v:$server /w:1024 /h:768" -PassThru
    $start = Get-Date

    # 3. Detect login via window title
    $loggedIn = $false
    while (((Get-Date) - $start).TotalSeconds -lt $TimeoutSec -and -not $proc.HasExited) {
        Start-Sleep -Seconds 2
        $title = (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue).MainWindowTitle
        if ($title -like "*$server*") { $loggedIn = $true; break }
    }

    if ($loggedIn) {
        $r.RdpAccess = 'YES'
        Write-Host "   Logged in. Session is live - SIGN OUT inside the remote session to continue..." -ForegroundColor Green

        # 4. Wait until the session ends (mstsc exits on logoff)
        $proc.WaitForExit()
        $dur = [int]((Get-Date) - $start).TotalSeconds
        $r.Reason = "Logged in; session closed after ${dur}s"
        Write-Host "   Session closed. Moving on." -ForegroundColor Green
    } else {
        $r.Reason = if ($proc.HasExited) { "mstsc closed before login (connection/credential error)" }
                    else { "No login within ${TimeoutSec}s (bad creds / no RDP logon right / NLA or policy block)" }
        if (-not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
        Write-Host "   RDP access: NO  $($r.Reason)" -ForegroundColor Red
    }

    cmdkey /delete:$target | Out-Null
    [pscustomobject]$r
}
}
finally {
    if ($null -eq $origAuthLevel) { Remove-ItemProperty $tsKey -Name AuthenticationLevelOverride -ErrorAction SilentlyContinue }
    else { Set-ItemProperty $tsKey -Name AuthenticationLevelOverride -Value $origAuthLevel -Type DWord }
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
