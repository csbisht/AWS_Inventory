<#
.SYNOPSIS
    Typing simulator: every N seconds, append a paragraph to a Notepad file, save,
    glide the mouse diagonally at random, then do the same for a Word document.
    Only types into those two windows: restores if minimised, brings to front,
    clicks inside, and verifies focus before pasting.
.PARAMETER NotesFile
    Notepad file name (saved on Desktop) or full path. ".txt" added if no extension.
.PARAMETER DocFile
    Word file name (saved on Desktop) or full path. ".docx" added if no extension.
.EXAMPLE
    .\Type-InNotepad.ps1 -Text "The cat sat on the mat." -NotesFile "Kids-notes" -DocFile "Kids-story" -IntervalSeconds 30
    (Count 0 = run until Ctrl+C in this PowerShell window)
	
	# run for 10 times or 10 lines
	Unblock-File .\Type-InNotepad-word_micemove.ps1
	powershell -ExecutionPolicy Bypass -File .\Type-InNotepad-word_micemove.ps1 -Text "The cat sat on the mat." -NotesFile "Kids-notes" -DocFile "Kids-story" -IntervalSeconds 30 -Count 10

.NOTES
    Requires Microsoft Word to open the .docx. Uses the clipboard (it is overwritten).
#>
param(
    [Parameter(Mandatory)][string]$Text,
    [string]$NotesFile = 'My-notes.txt',
    [string]$DocFile   = 'My-docs.docx',
    [int]$IntervalSeconds = 30,
    [int]$Count = 0
)

function Log($m) { Write-Host "[$(Get-Date -Format HH:mm:ss)] $m" }

if (-not ('TypeSimWin' -as [type])) {
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class TypeSimWin {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int InternalGetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int i);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint f, UIntPtr e);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }

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
}

$Desktop = [Environment]::GetFolderPath('Desktop')   # handles OneDrive Desktop
$wshell  = New-Object -ComObject WScript.Shell

function Resolve-TargetPath($f, $ext) {
    if (-not [IO.Path]::GetExtension($f)) { $f += $ext }
    if ([IO.Path]::IsPathRooted($f)) { return $f }
    return (Join-Path $Desktop $f)
}

$NotesPath = Resolve-TargetPath $NotesFile '.txt'
$DocPath   = Resolve-TargetPath $DocFile   '.docx'

$Targets = @(
    @{ Name = [IO.Path]::GetFileNameWithoutExtension($NotesPath); Path = $NotesPath; Proc = 'notepad'; Exe = 'notepad.exe'; Wait = 15 },
    @{ Name = [IO.Path]::GetFileNameWithoutExtension($DocPath);   Path = $DocPath;   Proc = 'WINWORD'; Exe = 'winword.exe'; Wait = 60 }
)

# Minimal valid empty .docx built without launching Word
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

Log "Notepad file: $NotesPath"
Log "Word file:    $DocPath"
if (-not (Test-Path $NotesPath)) { New-Item -Path $NotesPath -ItemType File -Force | Out-Null; Log "Created $NotesPath" }
if (-not (Test-Path $DocPath))   { New-EmptyDocx $DocPath; Log "Created $DocPath" }

function Find-TargetWindow($t) {
    $pids = @(Get-Process -Name $t.Proc -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    if ($pids.Count -eq 0) { return [IntPtr]::Zero }
    return [TypeSimWin]::Find($t.Name, [int[]]$pids)
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
    if ([TypeSimWin]::IsIconic($h)) { [TypeSimWin]::ShowWindow($h, 9) | Out-Null }   # SW_RESTORE
    for ($try = 0; $try -lt 5 -and [TypeSimWin]::GetForegroundWindow() -ne $h; $try++) {
        [TypeSimWin]::SetForegroundWindow($h) | Out-Null
        Start-Sleep -Milliseconds 250
        if ([TypeSimWin]::GetForegroundWindow() -ne $h) {
            [TypeSimWin]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)   # tap Alt to lift foreground lock
            [TypeSimWin]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
            [TypeSimWin]::SetForegroundWindow($h) | Out-Null
            Start-Sleep -Milliseconds 250
        }
    }
    if ([TypeSimWin]::GetForegroundWindow() -ne $h) { return $false }

    $r = New-Object TypeSimWin+RECT
    [TypeSimWin]::GetWindowRect($h, [ref]$r) | Out-Null
    [TypeSimWin]::SetCursorPos([int](($r.Left + $r.Right) / 2), [int]($r.Top + ($r.Bottom - $r.Top) * 0.6)) | Out-Null
    [TypeSimWin]::mouse_event(0x02, 0, 0, 0, [UIntPtr]::Zero)
    [TypeSimWin]::mouse_event(0x04, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 300

    return ([TypeSimWin]::GetForegroundWindow() -eq $h)
}

# Glide the pointer diagonally (random left/right, up/down), kept inside the screen(s)
function Move-MouseRandom {
    $p = New-Object TypeSimWin+POINT
    [TypeSimWin]::GetCursorPos([ref]$p) | Out-Null
    $vx = [TypeSimWin]::GetSystemMetrics(76); $vy = [TypeSimWin]::GetSystemMetrics(77)
    $vw = [TypeSimWin]::GetSystemMetrics(78); $vh = [TypeSimWin]::GetSystemMetrics(79)

    $d  = Get-Random -Minimum 150 -Maximum 400
    $tx = $p.X + $d * (Get-Random -InputObject @(-1, 1))
    $ty = $p.Y + $d * (Get-Random -InputObject @(-1, 1))
    $tx = [Math]::Min([Math]::Max($tx, $vx + 10), $vx + $vw - 10)
    $ty = [Math]::Min([Math]::Max($ty, $vy + 10), $vy + $vh - 10)

    $steps = 40
    for ($k = 1; $k -le $steps; $k++) {
        [TypeSimWin]::SetCursorPos([int]($p.X + ($tx - $p.X) * $k / $steps), [int]($p.Y + ($ty - $p.Y) * $k / $steps)) | Out-Null
        Start-Sleep -Milliseconds 15
    }
}

function Write-ToTarget($t, $i) {
    try {
        $h = Get-TargetWindow $t
        Log "#$i focusing $($t.Name)"
        if (Set-WindowFront $h) {
            $wshell.SendKeys('^{END}')
            Start-Sleep -Milliseconds 200
            if ($t.Proc -eq 'WINWORD') {
                Set-Clipboard -Value $Text          # Word drops trailing newline on paste
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
            Move-MouseRandom
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
