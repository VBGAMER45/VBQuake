# Starts QuakeVB visibly with a -keyfile, captures its window after 2.5 s and
# waits for it to quit (kills it after 15 s). Used to check the real window.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\capture_window.ps1 -exe <QuakeVB.exe> -keyfile <keys.txt> -out <shot.png>
# The keyfile must quit the game (` then quit ENTER) late enough for the capture.
param([string]$exe, [string]$keyfile, [string]$out)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public class W { [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern IntPtr FindWindow(string c, string t);
 public struct RECT { public int L, T, R, B; } }
"@
$p = Start-Process -FilePath $exe -ArgumentList "-keyfile $keyfile" -PassThru -WorkingDirectory (Split-Path $exe)
Start-Sleep -Milliseconds 2500
$p.Refresh()
$h = [W]::FindWindow("ThunderRT6FormDC", "QuakeVB")
[W]::SetForegroundWindow($h) | Out-Null
Start-Sleep -Milliseconds 400
$r = New-Object W+RECT
[W]::GetWindowRect($h, [ref]$r) | Out-Null
$bmp = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size)
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
"window $($r.R - $r.L) x $($r.B - $r.T), title '$($p.MainWindowTitle)'"
if (-not $p.WaitForExit(15000)) { "still running: killing"; $p.Kill() } else { "exited with code $($p.ExitCode)" }
