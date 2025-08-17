Add-Type -AssemblyName System.Windows.Forms

# Add user32.dll for mouse click
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
}
"@

$MOUSEEVENTF_LEFTDOWN = 0x02
$MOUSEEVENTF_LEFTUP   = 0x04

$rand = New-Object System.Random

# Anchor point = current mouse position
$anchor = [System.Windows.Forms.Cursor]::Position

# Define the "box" around the anchor in pixels
$boxSize = 50   # 50px wide/tall box (so ±25px from anchor)

while ($true) {
    # Pick random offset inside the box
    $offsetX = $rand.Next(-$boxSize/2, $boxSize/2)
    $offsetY = $rand.Next(-$boxSize/2, $boxSize/2)

    $x = $anchor.X + $offsetX
    $y = $anchor.Y + $offsetY

    # Move cursor
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)

    # Click
    [Mouse]::mouse_event($MOUSEEVENTF_LEFTDOWN,0,0,0,0)
    Start-Sleep -Milliseconds 50
    [Mouse]::mouse_event($MOUSEEVENTF_LEFTUP,0,0,0,0)

    Write-Host "Clicked at ($x,$y)"

    # Wait randomly 2–5 seconds before next click
    Start-Sleep -Seconds ($rand.Next(2,6))
}
