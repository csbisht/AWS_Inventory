<#
.SYNOPSIS
    Every N seconds: append a paragraph to Desktop\My-notes.txt (Notepad), save,
    then append the same paragraph to Desktop\My-docs.docx (Word), save.
    Only types into those two windows: restores if minimised, brings to front,
    clicks inside, and verifies focus before pasting.
.EXAMPLE
    .\Type-InNotepad.ps1 -Text "Hello world." -IntervalSeconds 30 -Count 10
    (Count 0 = run until Ctrl+C in this PowerShell window)
	powershell -ExecutionPolicy Bypass -File .\Type-InNotepad-word.ps1 -Text "Your paragraph here"
.NOTES
    Requires Microsoft Word to open the .docx. Uses the clipboard (it is overwritten).
#>
param(
    [Parameter(Mandatory)][string]$Text,
    [int]$IntervalSeconds = 30,
    [int]$Count = 0
)

function Log($m) { Write-Host "[$(Get-Date -Format HH:mm:ss)] $m" }

Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class Win {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    // InternalGetWindowText does NOT send messages, so a hung window can't block us
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int InternalGetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint f, UIntPtr e);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

    // Visible top-level window owned by one of 'pids' whose title contains 'part'
    public static IntPtr Find(string part, int[] pids) {
        var set = new HashSet<int>(pids);
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, l) => {
            if (!IsWindowVisible(h)) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!set.Contains((int)pid)) return true;
            var sb = new StringBuilder(512);
            InternalGetWindowText(h, sb, 512);
            if (sb.ToString().IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@

$Desktop = [Environment]::GetFolderPath('Desktop')   # handles OneDrive Desktop
$wshell  = New-Object -ComObject WScript.Shell

$Targets = @(
    @{ Name = 'My-notes'; Path = Join-Path $Desktop 'My-notes.txt'; Proc = 'notepad'; Exe = 'notepad.exe'; Wait = 15 },
    @{ Name = 'My-docs';  Path = Join-Path $Desktop 'My-docs.docx'; Proc = 'WINWORD'; Exe = 'winword.exe'; Wait = 60 }
)

# Build a minimal valid empty .docx without launching Word (no COM, nothing to hang)
function New-EmptyDocx($p) {
    Add-Type -AssemblyName System.IO.Compression
    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        '_rels/.rels'         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        'word/document.xml'   = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'
    }
    $enc = New-Object System.Text.UTF8Encoding($false)
    $fs  = [System.IO.File]::Open($p, [System.IO.FileMode]::Create)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    foreach ($k in $parts.Keys) {
        $sw = New-Object System.IO.StreamWriter($zip.CreateEntry($k).Open(), $enc)
        $sw.Write($parts[$k]); $sw.Dispose()
    }
    $zip.Dispose(); $fs.Dispose()
}

Log "Desktop: $Desktop"
if (-not (Test-Path $Targets[0].Path)) { New-Item -Path $Targets[0].Path -ItemType File | Out-Null; Log "Created $($Targets[0].Path)" }
if (-not (Test-Path $Targets[1].Path)) { New-EmptyDocx $Targets[1].Path; Log "Created $($Targets[1].Path)" }

function Find-TargetWindow($t) {
    $pids = @(Get-Process -Name $t.Proc -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    if ($pids.Count -eq 0) { return [IntPtr]::Zero }
    return [Win]::Find($t.Name, [int[]]$pids)
}

function Get-TargetWindow($t) {
    $h = Find-TargetWindow $t
    if ($h -ne [IntPtr]::Zero) { return $h }

    Log "Opening $($t.Path) ..."
    Start-Process $t.Exe -ArgumentList "`"$($t.Path)`""
    for ($s = 0; $s -lt ($t.Wait * 2); $s++) {
        Start-Sleep -Milliseconds 500
        $h = Find-TargetWindow $t
        if ($h -ne [IntPtr]::Zero) { Start-Sleep -Seconds 2; return $h }
    }
    throw "Timed out waiting for the window of $($t.Path)."
}

function Set-WindowFront([IntPtr]$h) {
    if ([Win]::IsIconic($h)) { [Win]::ShowWindow($h, 9) | Out-Null }   # SW_RESTORE
    for ($try = 0; $try -lt 5 -and [Win]::GetForegroundWindow() -ne $h; $try++) {
        [Win]::SetForegroundWindow($h) | Out-Null
        Start-Sleep -Milliseconds 250
        if ([Win]::GetForegroundWindow() -ne $h) {
            # Tap Alt to lift Windows' foreground lock, then retry
            [Win]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
            [Win]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
            [Win]::SetForegroundWindow($h) | Out-Null
            Start-Sleep -Milliseconds 250
        }
    }
    if ([Win]::GetForegroundWindow() -ne $h) { return $false }   # never click on another window

    # Click inside the document area (below title/ribbon/menu)
    $r = New-Object Win+RECT
    [Win]::GetWindowRect($h, [ref]$r) | Out-Null
    [Win]::SetCursorPos([int](($r.Left + $r.Right) / 2), [int]($r.Top + ($r.Bottom - $r.Top) * 0.6)) | Out-Null
    [Win]::mouse_event(0x02, 0, 0, 0, [UIntPtr]::Zero)
    [Win]::mouse_event(0x04, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 300

    return ([Win]::GetForegroundWindow() -eq $h)
}

function Write-ToTarget($t, $i) {
    try {
        $h = Get-TargetWindow $t
        Log "#$i focusing $($t.Name)"
        if (Set-WindowFront $h) {
            $wshell.SendKeys('^{END}')
            Start-Sleep -Milliseconds 200
            if ($t.Proc -eq 'WINWORD') {
                # Word drops a trailing newline on paste -> paste text, then press Enter
                Set-Clipboard -Value $Text
                $wshell.SendKeys('^v')
                Start-Sleep -Milliseconds 500
                $wshell.SendKeys('{ENTER}')
            } else {
                Set-Clipboard -Value ($Text + "`r`n")
                $wshell.SendKeys('^v')
            }
            Start-Sleep -Milliseconds 500
            $wshell.SendKeys('^s')
            Start-Sleep -Milliseconds 700
            Log "#$i wrote + saved -> $($t.Path)"
        } else {
            Write-Warning "#$i couldn't focus $($t.Name) window; skipped (nothing typed)."
        }
    } catch {
        Write-Warning "#$i $($t.Name): $_"
    }
}

$i = 0
while ($Count -eq 0 -or $i -lt $Count) {
    $i++
    foreach ($t in $Targets) { Write-ToTarget $t $i }   # Notepad first, then Word
    if ($Count -ne 0 -and $i -ge $Count) { break }
    Log "Sleeping $IntervalSeconds s (Ctrl+C to stop)"
    Start-Sleep -Seconds $IntervalSeconds
}
