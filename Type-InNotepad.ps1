<#
.SYNOPSIS
    Opens Notepad and pastes a given paragraph every N seconds.
.EXAMPLE
    .\Type-InNotepad.ps1 -Text "Hello world." -IntervalSeconds 30 -Count 10
    (Count 0 = run until Ctrl+C in this PowerShell window)
	powershell -ExecutionPolicy Bypass -File .\Type-InNotepad.ps1 -Text "Your paragraph here"
.NOTES
    Uses the clipboard to paste, so the clipboard content is overwritten.
#>
param(
    [Parameter(Mandatory)][string]$Text,
    [int]$IntervalSeconds = 30,
    [int]$Count = 0
)

$wshell = New-Object -ComObject WScript.Shell

Start-Process notepad.exe

# Wait until the Notepad window can actually be activated (max ~10s)
$ready = $false
for ($t = 0; $t -lt 20 -and -not $ready; $t++) {
    Start-Sleep -Milliseconds 500
    $ready = $wshell.AppActivate('Notepad')
}
if (-not $ready) { throw "Notepad window did not appear." }
Start-Sleep -Seconds 2   # let the editor finish initialising

$i = 0
while ($Count -eq 0 -or $i -lt $Count) {
    $i++
    if ($wshell.AppActivate('Notepad')) {
        Start-Sleep -Milliseconds 400
        Set-Clipboard -Value ($Text + "`r`n`r`n")
        $wshell.SendKeys('^{END}')      # caret to end of document
        Start-Sleep -Milliseconds 150
        $wshell.SendKeys('^v')          # paste atomically
        Write-Host "[$(Get-Date -Format HH:mm:ss)] Pasted paragraph #$i"
    } else {
        Write-Warning "Notepad window not found; skipping iteration $i"
    }
    if ($Count -ne 0 -and $i -ge $Count) { break }
    Start-Sleep -Seconds $IntervalSeconds
}
