<#
.SYNOPSIS
    Tests whether a user can RDP into multiple Windows servers and reports
    which servers grant access and which don't.

.EXAMPLE
    .\Test-RdpAccess.ps1 -ServerFile .\servers.txt
.EXAMPLE
    .\Test-RdpAccess.ps1 -Servers srv01,srv02 -Mode Interactive
#>
[CmdletBinding()]
param(
    [string[]]$Servers,
    [string]$ServerFile,
    [ValidateSet('Port','Auth','Interactive')]
    [string]$Mode = 'Auth',
    [switch]$CheckGroup,
    [string]$FreeRdpPath = 'wfreerdp.exe',
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

if ($Mode -eq 'Auth' -and -not (Get-Command $FreeRdpPath -ErrorAction SilentlyContinue)) {
    throw "FreeRDP not found ($FreeRdpPath). Install it ('winget install FreeRDP.FreeRDP') or use -Mode Interactive."
}

function Get-CredForServer([string]$server) {
    $u = if ($User -like '.\*') { "$($server.Split('.')[0])\$($User.Substring(2))" } else { $User }
    New-Object pscredential ($u, (ConvertTo-SecureString $Pass -AsPlainText -Force))
}

$results = foreach ($server in $Servers) {
    $server = $server.Trim()
    $cred   = Get-CredForServer $server
    $nc     = $cred.GetNetworkCredential()
    Write-Host "`n==> $server  (as $($cred.UserName))" -ForegroundColor Cyan

    $r = [ordered]@{
        Server    = $server
        User      = $cred.UserName
        PortOpen  = $false
        RdpAccess = 'NO'
        RdpGroup  = 'N/A'
        Reason    = ''
    }

    # 1. Port check
    try {
        $r.PortOpen = [bool](Test-NetConnection -ComputerName $server -Port $Port -WarningAction SilentlyContinue).TcpTestSucceeded
    } catch { $r.Reason = "Port test error: $($_.Exception.Message)" }

    if (-not $r.PortOpen) {
        if (-not $r.Reason) { $r.Reason = "Port $Port closed/unreachable" }
        Write-Host "   $($r.Reason)" -ForegroundColor Red
        [pscustomobject]$r; continue
    }

    # 2. Optional group membership check (needs WinRM + admin on target)
    if ($CheckGroup) {
        try {
            $r.RdpGroup = Invoke-Command -ComputerName $server -Credential $cred -ErrorAction Stop -ScriptBlock {
                (Get-LocalGroupMember -Group 'Remote Desktop Users').Name -join '; '
            }
        } catch { $r.RdpGroup = "Error: $($_.Exception.Message)" }
    }

    # 3. Access test
    switch ($Mode) {
        'Port' {
            $r.RdpAccess = 'PORT-ONLY'
            $r.Reason    = 'Port open; credentials not tested'
        }
        'Auth' {
            $frdpArgs = @("/v:$server", "/port:$Port", "/u:$($nc.UserName)", "/p:$Pass", '/cert:ignore', '+auth-only')
            if ($nc.Domain) { $frdpArgs += "/d:$($nc.Domain)" }

            $p = Start-Process -FilePath $FreeRdpPath -ArgumentList $frdpArgs -Wait -PassThru -NoNewWindow `
                 -RedirectStandardOutput "$env:TEMP\frdp_out.txt" -RedirectStandardError "$env:TEMP\frdp_err.txt"
            if ($p.ExitCode -eq 0) {
                $r.RdpAccess = 'YES'
            } else {
                $r.Reason = "Auth failed (exit $($p.ExitCode)): " + ((Get-Content "$env:TEMP\frdp_err.txt" -Tail 2) -join ' | ')
            }
        }
        'Interactive' {
            $target = "TERMSRV/$server"
            cmdkey /generic:$target /user:$($cred.UserName) /pass:$Pass | Out-Null
            Start-Process mstsc -ArgumentList "/v:$server" -Wait
            cmdkey /delete:$target | Out-Null
            if ((Read-Host "   Did login succeed on $server? (y/n)") -match '^y') {
                $r.RdpAccess = 'YES'
            } else {
                $r.Reason = Read-Host "   Error seen (optional)"
            }
        }
    }

    $color = if ($r.RdpAccess -eq 'YES') { 'Green' } else { 'Red' }
    Write-Host "   RDP access: $($r.RdpAccess) $($r.Reason)" -ForegroundColor $color
    [pscustomobject]$r
}

# ============================== SUMMARY ==============================
$yes = $results | Where-Object RdpAccess -eq 'YES'
$no  = $results | Where-Object RdpAccess -ne 'YES'

Write-Host "`n========== RDP ACCESS SUMMARY ==========" -ForegroundColor Yellow
Write-Host "HAS ACCESS ($($yes.Count)):" -ForegroundColor Green
$yes | ForEach-Object { Write-Host "  [YES] $($_.Server)" -ForegroundColor Green }
Write-Host "NO ACCESS ($($no.Count)):" -ForegroundColor Red
$no  | ForEach-Object { Write-Host "  [NO]  $($_.Server)  -> $($_.Reason)" -ForegroundColor Red }

$results | Format-Table Server, User, PortOpen, RdpAccess, Reason -AutoSize
$results | Export-Csv -Path $OutCsv -NoTypeInformation
Write-Host "Results saved to $OutCsv" -ForegroundColor Yellow
