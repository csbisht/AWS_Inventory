<#
.SYNOPSIS
    CLI-only RDP access check for multiple Windows servers. No GUI, no session created.

.DESCRIPTION
    Per server:
      1. TCP 3389 reachable
      2. RDP enabled on server (fDenyTSConnections = 0)
      3. Password valid for the test user
      4. User's effective groups (incl. nested) hold SeRemoteInteractiveLogonRight
         ("Allow log on through RDS") and NOT SeDenyRemoteInteractiveLogonRight
    Steps 2 and 4 run on the target via WinRM as the account EXECUTING this script
    (needs local admin on target + WinRM enabled).

.EXAMPLE
    .\Test-RdpAccess.ps1 -ServerFile .\servers.txt
.EXAMPLE
    .\Test-RdpAccess.ps1 -Servers srv01,srv02
#>
[CmdletBinding()]
param(
    [string[]]$Servers,
    [string]$ServerFile,
    [int]$Port = 3389,
    [string]$OutCsv = ".\rdp-test-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
)

# ====================== HARDCODED TEST CREDENTIALS ======================
# Domain account : "DOMAIN\testuser" or "testuser@corp.example.com"
# Local account  : ".\localuser"  (checked as a local user on each server)
$User = "DOMAIN\testuser"
$Pass = "P@ssw0rd!"
# ========================================================================

if ($ServerFile) { $Servers += Get-Content $ServerFile | Where-Object { $_.Trim() -and $_ -notmatch '^\s*#' } }
if (-not $Servers) { throw "No servers supplied. Use -Servers or -ServerFile." }

Add-Type -AssemblyName System.DirectoryServices.AccountManagement

$IsLocal = $User -like '.\*'
$Sam = $User
if ($Sam -like '*\*') { $Sam = $Sam.Split('\')[1] }
if ($Sam -like '*@*') { $Sam = $Sam.Split('@')[0] }

# Runs ON the target server
$remoteCheck = {
    param($sam, $isLocal)
    $o = [ordered]@{ RdpEnabled = $false; Allowed = $false; Detail = '' }

    $ts = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    $o.RdpEnabled = ($ts.fDenyTSConnections -eq 0)

    # --- Effective SIDs of the user (user + all groups, nested, incl. local groups) ---
    if ($isLocal) {
        $u = Get-LocalUser -Name $sam -ErrorAction Stop
        if (-not $u.Enabled) { $o.Detail = 'Local account disabled'; return [pscustomobject]$o }
        $sids = @($u.SID.Value, 'S-1-1-0', 'S-1-5-11', 'S-1-5-32-545')
        foreach ($g in Get-LocalGroup) {
            if (Get-LocalGroupMember $g -ErrorAction SilentlyContinue | Where-Object { $_.SID -eq $u.SID }) { $sids += $g.SID.Value }
        }
    } else {
        $res = ([adsisearcher]"(&(objectCategory=user)(sAMAccountName=$sam))").FindOne()
        if (-not $res) { $o.Detail = "User $sam not found in AD"; return [pscustomobject]$o }
        $upn = $res.Properties['userprincipalname'] | Select-Object -First 1
        if (-not $upn) { $o.Detail = "User $sam has no UPN; can't build token"; return [pscustomobject]$o }
        $id = [Security.Principal.WindowsIdentity]::new($upn)   # S4U token, includes nested + local groups
        $sids = @($id.User.Value) + @($id.Groups | ForEach-Object { $_.Value })
    }

    # --- User rights on this server ---
    $tmp = Join-Path $env:TEMP "ur_$([guid]::NewGuid()).inf"
    secedit /export /areas USER_RIGHTS /cfg $tmp | Out-Null
    $cfg = Get-Content $tmp
    Remove-Item $tmp -Force

    function Get-RightSids($name) {
        $line = $cfg | Where-Object { $_ -match "^$name\s*=" }
        if (-not $line) { return @() }
        ($line -replace "^$name\s*=\s*", '') -split ',' | ForEach-Object {
            $e = $_.Trim()
            if ($e.StartsWith('*')) { $e.TrimStart('*') }
            else { try { ([Security.Principal.NTAccount]$e).Translate([Security.Principal.SecurityIdentifier]).Value } catch { $e } }
        }
    }
    function Get-Name($sid) {
        try { ([Security.Principal.SecurityIdentifier]$sid).Translate([Security.Principal.NTAccount]).Value } catch { $sid }
    }

    $allowHit = @(Get-RightSids 'SeRemoteInteractiveLogonRight'     | Where-Object { $_ -in $sids })
    $denyHit  = @(Get-RightSids 'SeDenyRemoteInteractiveLogonRight' | Where-Object { $_ -in $sids })

    if ($denyHit)      { $o.Detail = "DENIED via: " + (($denyHit | ForEach-Object { Get-Name $_ }) -join ', ') }
    elseif ($allowHit) { $o.Allowed = $true; $o.Detail = "Allowed via: " + (($allowHit | ForEach-Object { Get-Name $_ }) -join ', ') }
    else               { $o.Detail = "No 'Allow log on through RDS' right (not in Remote Desktop Users/Administrators)" }

    [pscustomobject]$o
}

$results = foreach ($server in $Servers) {
    $server = $server.Trim()
    Write-Host "`n==> $server" -ForegroundColor Cyan
    $r = [ordered]@{ Server = $server; User = $User; PortOpen = $false; RdpEnabled = $null; CredValid = $null; RdpAccess = 'NO'; Reason = '' }

    # 1. Port
    try { $r.PortOpen = [bool](Test-NetConnection $server -Port $Port -WarningAction SilentlyContinue).TcpTestSucceeded } catch {}
    if (-not $r.PortOpen) { $r.Reason = "Port $Port closed/unreachable" }

    # 2. Password
    if ($r.PortOpen) {
        try {
            $ctx = if ($IsLocal) { New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine', $server) }
                   else          { New-Object DirectoryServices.AccountManagement.PrincipalContext('Domain') }
            $r.CredValid = $ctx.ValidateCredentials($Sam, $Pass)
            if (-not $r.CredValid) { $r.Reason = 'Invalid username/password (or account locked/disabled)' }
        } catch { $r.CredValid = $false; $r.Reason = "Cred check error: $($_.Exception.Message)" }
    }

    # 3. RDP enabled + logon right (on target)
    if ($r.CredValid) {
        try {
            $rc = Invoke-Command -ComputerName $server -ScriptBlock $remoteCheck -ArgumentList $Sam, $IsLocal -ErrorAction Stop
            $r.RdpEnabled = $rc.RdpEnabled
            if (-not $rc.RdpEnabled)   { $r.Reason = 'RDP disabled on server (fDenyTSConnections=1)' }
            elseif ($rc.Allowed)       { $r.RdpAccess = 'YES'; $r.Reason = $rc.Detail }
            else                       { $r.Reason = $rc.Detail }
        } catch { $r.Reason = "Remote check failed (WinRM/admin?): $($_.Exception.Message)" }
    }

    $color = if ($r.RdpAccess -eq 'YES') { 'Green' } else { 'Red' }
    Write-Host "   RDP access: $($r.RdpAccess)  $($r.Reason)" -ForegroundColor $color
    [pscustomobject]$r
}

# ============================== SUMMARY ==============================
$yes = @($results | Where-Object RdpAccess -eq 'YES')
$no  = @($results | Where-Object RdpAccess -ne 'YES')

Write-Host "`n========== RDP ACCESS SUMMARY ==========" -ForegroundColor Yellow
Write-Host "HAS ACCESS ($($yes.Count)):" -ForegroundColor Green
$yes | ForEach-Object { Write-Host "  [YES] $($_.Server)  -> $($_.Reason)" -ForegroundColor Green }
Write-Host "NO ACCESS ($($no.Count)):" -ForegroundColor Red
$no  | ForEach-Object { Write-Host "  [NO]  $($_.Server)  -> $($_.Reason)" -ForegroundColor Red }

$results | Format-Table Server, PortOpen, RdpEnabled, CredValid, RdpAccess, Reason -AutoSize
$results | Export-Csv -Path $OutCsv -NoTypeInformation
Write-Host "Results saved to $OutCsv" -ForegroundColor Yellow
