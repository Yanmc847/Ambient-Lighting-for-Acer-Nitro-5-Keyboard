$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

$null = Add-Type -AssemblyName System.Drawing
$null = Add-Type -AssemblyName System.Windows.Forms
$null = Add-Type -AssemblyName System.Web.Extensions

[System.Windows.Forms.MessageBox]::Show( "Started", "AmbientKeyboard", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information ) | Out-Null

$configFile = "$env:APPDATA\AmbientKeyboardConfig.json"
if (Test-Path $configFile) {
    $config = Get-Content $configFile | ConvertFrom-Json
} else {
    $config = [PSCustomObject]@{
        TimerInterval = 33
        MaxLum = 80
        ShowRGB = $false
        SkipThreshold = 0
    }
}

$global:lastColors = @(0,0,0,0)
$global:showRGB = $config.ShowRGB
$global:timerInterval = $config.TimerInterval
$global:maxLum = $config.MaxLum
$global:skipThreshold = $config.SkipThreshold


Add-Type @"
using System; 
using System.Runtime.InteropServices;
using System.Drawing;

public class DPI {  
    [DllImport("gdi32.dll")]
    static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    public enum DeviceCap {
        VERTRES = 10,
        DESKTOPVERTRES = 117
    } 

    public static float scaling() {
        Graphics g = Graphics.FromHwnd(IntPtr.Zero);
        IntPtr desktop = g.GetHdc();
        int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
        int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);
        return (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
    }
}
"@ -ReferencedAssemblies 'System.Drawing.dll' -ErrorAction Stop

$ScaleFactor = [DPI]::scaling()

$screenWidth  = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$zoneWidth    = [math]::Floor(($screenWidth * $ScaleFactor) / 4)
$zoneNumbers  = @(1, 2, 4, 8)

try { $wmiObj = Get-WmiObject -Namespace root\wmi -Class AcerGamingFunction -ErrorAction Stop }
catch { $wmiObj = $null }

function Get-AverageColorSample($bitmap, $rect, $step=10) {
    $totalR=0; $totalG=0; $totalB=0; $count=0
    try {
        for ($x=$rect.X; $x -lt ($rect.X+$rect.Width); $x+=$step) {
            for ($y=$rect.Y; $y -lt ($rect.Y+$rect.Height); $y+=$step) {
                $color = $bitmap.GetPixel([int]$x,[int]$y)
                $totalR += $color.R; $totalG += $color.G; $totalB += $color.B; $count++
            }
        }
    } catch {}
    if ($count -eq 0) { return [System.Drawing.Color]::FromArgb(0,0,0) }
    return [System.Drawing.Color]::FromArgb([int]($totalR/$count),[int]($totalG/$count),[int]($totalB/$count))
}

function Convert-ColorToAcerNumber($color, $zoneNumber) {
    $hex = "{0:X2}{1:X2}{2:X2}00" -f $color.B,$color.G,$color.R
    return [convert]::ToUInt32($hex,16) + $zoneNumber
}

function ColorDifference($c1,$c2) {
    return [math]::Abs($c1.R-$c2.R)+[math]::Abs($c1.G-$c2.G)+[math]::Abs($c1.B-$c2.B)
}

$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)) { exit 0 }

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
$notifyIcon.Text = "RGB Keyboard"
$notifyIcon.Visible = $true

$form = New-Object System.Windows.Forms.Form
$form.Text = "AmbientKeyboard Settings"
$form.Size = New-Object System.Drawing.Size(350,210)

$labelFPS = New-Object System.Windows.Forms.Label
$labelFPS.Text = "FPS:"; $labelFPS.Location = '10,10'; $form.Controls.Add($labelFPS)
$numericFPS = New-Object System.Windows.Forms.NumericUpDown
$numericFPS.Minimum = 1; $numericFPS.Maximum = 120
$numericFPS.Value = [math]::Round(1000 / $global:timerInterval)
$numericFPS.Location = '120,10'; $form.Controls.Add($numericFPS)

$labelMaxLum = New-Object System.Windows.Forms.Label
$labelMaxLum.Text="Maximum brightness:";$labelMaxLum.Location='10,40'; $form.Controls.Add($labelMaxLum)
$numericMaxLum = New-Object System.Windows.Forms.NumericUpDown
$numericMaxLum.Minimum=0;$numericMaxLum.Maximum=100;$numericMaxLum.Value=$global:maxLum;$numericMaxLum.Location='160,40'
$form.Controls.Add($numericMaxLum)

$labelSkip = New-Object System.Windows.Forms.Label
$labelSkip.Text="Skip threshold:";$labelSkip.Location='10,70'; $form.Controls.Add($labelSkip)
$numericSkip = New-Object System.Windows.Forms.NumericUpDown
$numericSkip.Minimum=0;$numericSkip.Maximum=765;$numericSkip.Value=$global:skipThreshold;$numericSkip.Location='160,70'
$form.Controls.Add($numericSkip)

$checkRGB = New-Object System.Windows.Forms.CheckBox
$checkRGB.Text="Show RGB values";$checkRGB.Location='10,100';$checkRGB.Checked=$global:showRGB
$form.Controls.Add($checkRGB)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text="Apply";$btnApply.Location='10,130'
$btnApply.Add_Click({
    $fps = [int]$numericFPS.Value; if ($fps -le 0) { $fps = 1 }
    $global:timerInterval = [math]::Round(1000 / $fps); $timer.Interval = $global:timerInterval
    $global:maxLum = [int]$numericMaxLum.Value
    $global:showRGB = $checkRGB.Checked
    $global:skipThreshold = [int]$numericSkip.Value

    $config | Add-Member -Force TimerInterval $global:timerInterval
    $config | Add-Member -Force MaxLum $global:maxLum
    $config | Add-Member -Force ShowRGB $global:showRGB
    $config | Add-Member -Force SkipThreshold $global:skipThreshold
    $null = $config | ConvertTo-Json | Set-Content $configFile

    if ($global:maxLum -ne $global:lastLum) {
        $backlightArray = [byte[]](0,0,[byte]$global:maxLum,255,0,0,0,0,75,0,0,0,0,0,0,0)
        if ($wmiObj) { $null = $wmiObj.SetGamingKBBacklight($backlightArray) }
        $global:lastLum = $global:maxLum
    }
})
$form.Controls.Add($btnApply)

$menu = New-Object System.Windows.Forms.ContextMenu
$openItem = New-Object System.Windows.Forms.MenuItem "Settings"; $openItem.Add_Click({ $form.ShowDialog() })
$exitItem = New-Object System.Windows.Forms.MenuItem "Exit"; $exitItem.Add_Click({ $notifyIcon.Visible=$false; [System.Windows.Forms.Application]::Exit() })
$menu.MenuItems.AddRange(@($openItem, $exitItem))
$notifyIcon.ContextMenu = $menu

$overlayForm = New-Object System.Windows.Forms.Form
$overlayForm.FormBorderStyle='None'
$overlayForm.BackColor=[System.Drawing.Color]::Black
$overlayForm.TransparencyKey=[System.Drawing.Color]::Black
$overlayForm.TopMost=$true
$overlayForm.ShowInTaskbar=$false
$overlayForm.StartPosition='Manual'
$overlayForm.Size=New-Object System.Drawing.Size($screenWidth,$screenHeight)
$overlayForm.Location=New-Object System.Drawing.Point(0,0)
$overlayForm.Visible=$true

$zoneLabels=@(); $skipLabels=@()
for ($i=0;$i -lt 4;$i++){
    $label=New-Object System.Windows.Forms.Label
    $label.ForeColor=[System.Drawing.Color]::White
    $label.BackColor=[System.Drawing.Color]::FromArgb(200,0,0,0)
    $label.AutoSize=$true
    $label.Font=New-Object System.Drawing.Font("Arial",16,[System.Drawing.FontStyle]::Bold)
    $label.Location=New-Object System.Drawing.Point([int]($i*$zoneWidth+$zoneWidth/2-50),10)
    $overlayForm.Controls.Add($label);$zoneLabels+=$label

    $skip=New-Object System.Windows.Forms.Label
    $skip.ForeColor=[System.Drawing.Color]::Yellow
    $skip.BackColor=[System.Drawing.Color]::FromArgb(200,0,0,0)
    $skip.AutoSize=$true
    $skip.Font=New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Italic)
    $skip.Location=New-Object System.Drawing.Point([int]($i*$zoneWidth+$zoneWidth/2-25),35)
    $overlayForm.Controls.Add($skip);$skipLabels+=$skip
}
if ($wmiObj) { $null = $wmiObj.SetGamingKBBacklight([byte[]](0,0,[byte]$global:maxLum,255,0,0,0,0,75,0,0,0,0,0,0,0)) }
$global:lastLum = $global:maxLum

$timer=New-Object System.Windows.Forms.Timer
$timer.Interval=$global:timerInterval
$timer.Add_Tick({
    if (-not $wmiObj) { return }
    try {
        $bitmap=New-Object System.Drawing.Bitmap $screenWidth,$screenHeight
        $graphics=[System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen(0,0,0,0,$bitmap.Size)
        $currentColors=@()
        for($i=0;$i -lt 4;$i++){
            $rect=[System.Drawing.Rectangle]::new([int]($i*$zoneWidth),0,[int]$zoneWidth,[int]$screenHeight)
            $avgColor=Get-AverageColorSample $bitmap $rect
            $currentColors+=$avgColor
            $diff=ColorDifference $avgColor $global:lastColors[$i]
            if($diff -gt $global:skipThreshold){
                $null = $wmiObj.SetGamingRgbKb([uint32](Convert-ColorToAcerNumber $avgColor $zoneNumbers[$i]))
                $skipLabels[$i].Text=""
            } else { $skipLabels[$i].Text="Skip" }
if($global:showRGB){
    $zoneLabels[$i].Text = "R:$([math]::Round($avgColor.R / $global:ScaleFactor)) G:$([math]::Round($avgColor.G / $global:ScaleFactor)) B:$([math]::Round($avgColor.B / $global:ScaleFactor))"
    
    $zoneLabels[$i].Location = New-Object System.Drawing.Point(
        [int](($i * $global:zoneWidth + $global:zoneWidth/2 - 50) / $global:ScaleFactor),
        $zoneLabels[$i].Location.Y
    )

    $skipLabels[$i].Location = New-Object System.Drawing.Point(
        [int](($i * $global:zoneWidth + $global:zoneWidth/2 - 25) / $global:ScaleFactor),
        $skipLabels[$i].Location.Y
    )
} else {
    $zoneLabels[$i].Text = ""
    $skipLabels[$i].Text = ""
}


        }
        $global:lastColors=$currentColors
        $graphics.Dispose();$bitmap.Dispose()
    } catch {}
})
$timer.Start()

[System.Windows.Forms.Application]::Run()
exit 0