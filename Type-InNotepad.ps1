<#
.SYNOPSIS
    Appends a paragraph to Desktop\My-notes.txt via Notepad every N seconds, then saves (Ctrl+S).
    Only ever types into the My-notes.txt Notepad window: restores it if minimised,
    brings it to front, clicks inside it, and verifies focus before pasting.
.EXAMPLE
    .\Type-InNotepad.ps1 -Text "Hello world." -IntervalSeconds 30 -Count 10
    (Count 0 = run until Ctrl+C in this PowerShell window)
	powershell -ExecutionPolicy Bypass -File .\Type-InNotepad.ps1 -Text "Your paragraph here"
.NOTES
    Uses the clipboard to paste, so clipboard content is overwritten.
#>
param(
    [Parameter(Mandatory)][string]$Text,
    [int]$IntervalSeconds = 30,
    [int]$Count = 0
)

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint f, UIntPtr e);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

    public static IntPtr Find(string part) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, l) => {
            if (!IsWindowVisible(h)) return true;
            var sb = new StringBuilder(512);
            GetWindowText(h, sb, 512);
            string t = sb.ToString();
            if (t.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0 && t.EndsWith("Notepad")) {
                found = h; return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@

$FileName = 'My-notes.txt'
$Path     = Join-Path ([Environment]::GetFolderPath('Desktop')) $FileName   # handles OneDrive Desktop
$wshell   = New-Object -ComObject WScript.Shell

# 1) Create the file if missing
if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType File | Out-Null }

function Get-NotesWindow {
    $h = [Win]::Find($FileName)
    if ($h -ne [IntPtr]::Zero) { return $h }

    # Not open -> launch Notepad with the file and wait for its window (max ~15s)
    Start-Process notepad.exe -ArgumentList "`"$Path`""
    for ($t = 0; $t -lt 30; $t++) {
        Start-Sleep -Milliseconds 500
        $h = [Win]::Find($FileName)
        if ($h -ne [IntPtr]::Zero) { Start-Sleep -Seconds 2; return $h }
    }
    throw "Could not find the Notepad window for $FileName."
}

function Set-WindowFront([IntPtr]$h) {
    if ([Win]::IsIconic($h)) { [Win]::ShowWindow($h, 9) | Out-Null }   # SW_RESTORE
    for ($try = 0; $try -lt 5; $try++) {
        # Tap Alt to lift Windows' foreground lock, then activate
        [Win]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
        [Win]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
        [Win]::SetForegroundWindow($h) | Out-Null
        Start-Sleep -Milliseconds 300
        if ([Win]::GetForegroundWindow() -eq $h) { break }
    }

    # Click inside the editor area (centre-ish, below title/tab/menu bars)
    $r = New-Object Win+RECT
    [Win]::GetWindowRect($h, [ref]$r) | Out-Null
    $x = [int](($r.Left + $r.Right) / 2)
    $y = [int]($r.Top + ($r.Bottom - $r.Top) * 0.6)
    [Win]::SetCursorPos($x, $y) | Out-Null
    [Win]::mouse_event(0x02, 0, 0, 0, [UIntPtr]::Zero)   # left down
    [Win]::mouse_event(0x04, 0, 0, 0, [UIntPtr]::Zero)   # left up
    Start-Sleep -Milliseconds 300

    return ([Win]::GetForegroundWindow() -eq $h)
}

$i = 0
while ($Count -eq 0 -or $i -lt $Count) {
    $i++
    try {
        $h = Get-NotesWindow
        if (Set-WindowFront $h) {
            Set-Clipboard -Value ($Text + "`r`n`r`n")
            $wshell.SendKeys('^{END}')   # caret to end
            Start-Sleep -Milliseconds 150
            $wshell.SendKeys('^v')       # 2) paste into My-notes.txt only
            Start-Sleep -Milliseconds 300
            $wshell.SendKeys('^s')       # 3) save
            Write-Host "[$(Get-Date -Format HH:mm:ss)] Wrote + saved paragraph #$i -> $Path"
        } else {
            Write-Warning "Couldn't focus $FileName window; skipped iteration $i (nothing typed)."
        }
    } catch {
        Write-Warning $_
    }
    if ($Count -ne 0 -and $i -ge $Count) { break }
    Start-Sleep -Seconds $IntervalSeconds
}
