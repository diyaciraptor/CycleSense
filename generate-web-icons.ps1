Add-Type -AssemblyName System.Drawing
$sourcePath = 'D:\Period Tracker\Cycle Sense.png'
$source = [System.Drawing.Image]::FromFile($sourcePath)
$targets = @(
  @{ Path = 'D:\Period Tracker\icons\cyclesense-192.png'; Width = 192; Height = 192 },
  @{ Path = 'D:\Period Tracker\icons\cyclesense-512.png'; Width = 512; Height = 512 },
  @{ Path = 'D:\Period Tracker\icons\apple-touch-icon.png'; Width = 180; Height = 180 }
)
foreach ($target in $targets) {
  $bmp = [System.Drawing.Bitmap]::new($target.Width, $target.Height)
  $gfx = [System.Drawing.Graphics]::FromImage($bmp)
  $gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $gfx.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $gfx.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $gfx.DrawImage($source, 0, 0, $target.Width, $target.Height)
  $bmp.Save($target.Path, [System.Drawing.Imaging.ImageFormat]::Png)
  $gfx.Dispose()
  $bmp.Dispose()
}
$source.Dispose()