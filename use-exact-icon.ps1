Add-Type -AssemblyName System.Drawing

$sourcePath = 'D:\Period Tracker\Cycle Sense.png'
$base = 'D:\Period Tracker\android\app\src\main\res'
$source = [System.Drawing.Image]::FromFile($sourcePath)
$bitmapSource = [System.Drawing.Bitmap]::FromFile($sourcePath)
$topLeft = $bitmapSource.GetPixel(0,0)
$bgHex = '#{0:X2}{1:X2}{2:X2}' -f $topLeft.R, $topLeft.G, $topLeft.B

function Save-ResizedImage([System.Drawing.Image]$sourceImage, [string]$destPath, [int]$width, [int]$height) {
  $bmp = [System.Drawing.Bitmap]::new($width, $height)
  $gfx = [System.Drawing.Graphics]::FromImage($bmp)
  $gfx.Clear([System.Drawing.Color]::Transparent)
  $gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $gfx.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $gfx.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $gfx.DrawImage($sourceImage, 0, 0, $width, $height)
  $tempPath = "$destPath.tmp.png"
  if (Test-Path $tempPath) { Remove-Item -Force $tempPath }
  $bmp.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
  $gfx.Dispose()
  $bmp.Dispose()
  if (Test-Path $destPath) { Remove-Item -Force $destPath }
  Move-Item -Force $tempPath $destPath
}

$iconSizes = @{
  'mipmap-mdpi\ic_launcher.png' = @{ Width = 48; Height = 48 };
  'mipmap-hdpi\ic_launcher.png' = @{ Width = 72; Height = 72 };
  'mipmap-xhdpi\ic_launcher.png' = @{ Width = 96; Height = 96 };
  'mipmap-xxhdpi\ic_launcher.png' = @{ Width = 144; Height = 144 };
  'mipmap-xxxhdpi\ic_launcher.png' = @{ Width = 192; Height = 192 };
  'mipmap-mdpi\ic_launcher_round.png' = @{ Width = 48; Height = 48 };
  'mipmap-hdpi\ic_launcher_round.png' = @{ Width = 72; Height = 72 };
  'mipmap-xhdpi\ic_launcher_round.png' = @{ Width = 96; Height = 96 };
  'mipmap-xxhdpi\ic_launcher_round.png' = @{ Width = 144; Height = 144 };
  'mipmap-xxxhdpi\ic_launcher_round.png' = @{ Width = 192; Height = 192 };
  'mipmap-mdpi\ic_launcher_foreground.png' = @{ Width = 108; Height = 108 };
  'mipmap-hdpi\ic_launcher_foreground.png' = @{ Width = 162; Height = 162 };
  'mipmap-xhdpi\ic_launcher_foreground.png' = @{ Width = 216; Height = 216 };
  'mipmap-xxhdpi\ic_launcher_foreground.png' = @{ Width = 324; Height = 324 };
  'mipmap-xxxhdpi\ic_launcher_foreground.png' = @{ Width = 432; Height = 432 };
}

foreach ($entry in $iconSizes.GetEnumerator()) {
  Save-ResizedImage $source (Join-Path $base $entry.Key) $entry.Value.Width $entry.Value.Height
}

$splashFiles = Get-ChildItem -Path $base -Recurse -Filter 'splash.png'
foreach ($file in $splashFiles) {
  $existing = [System.Drawing.Image]::FromFile($file.FullName)
  $width = $existing.Width
  $height = $existing.Height
  $existing.Dispose()
  Save-ResizedImage $source $file.FullName $width $height
}

$bgPath = 'D:\Period Tracker\android\app\src\main\res\values\ic_launcher_background.xml'
$bgContent = @"
<?xml version="1.0" encoding="utf-8"?>
<resources>
    <color name="ic_launcher_background">$bgHex</color>
</resources>
"@
[System.IO.File]::WriteAllText($bgPath, $bgContent, (New-Object System.Text.UTF8Encoding($false)))

$bitmapSource.Dispose()
$source.Dispose()
Write-Output $bgHex