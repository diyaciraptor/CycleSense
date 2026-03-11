Add-Type -AssemblyName System.Drawing

function New-Brush($hex) {
  return [System.Drawing.SolidBrush]::new([System.Drawing.ColorTranslator]::FromHtml($hex))
}

function New-Pen($hex, $width) {
  $pen = [System.Drawing.Pen]::new([System.Drawing.ColorTranslator]::FromHtml($hex), $width)
  $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
  $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
  $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
  return $pen
}

function Draw-Hand([System.Drawing.Graphics]$g, [float]$cx, [float]$cy, [float]$scale, [bool]$flip, [System.Drawing.Brush]$fill, [System.Drawing.Pen]$outline, [System.Drawing.Pen]$accent) {
  $points = @(
    [System.Drawing.PointF]::new(-8, 18),
    [System.Drawing.PointF]::new(-13, 8),
    [System.Drawing.PointF]::new(-10, -6),
    [System.Drawing.PointF]::new(-5, -14),
    [System.Drawing.PointF]::new(-1, -11),
    [System.Drawing.PointF]::new(0, -28),
    [System.Drawing.PointF]::new(5, -31),
    [System.Drawing.PointF]::new(9, -28),
    [System.Drawing.PointF]::new(10, -12),
    [System.Drawing.PointF]::new(16, -9),
    [System.Drawing.PointF]::new(18, -2),
    [System.Drawing.PointF]::new(18, 8),
    [System.Drawing.PointF]::new(11, 18),
    [System.Drawing.PointF]::new(3, 22)
  )
  $mapped = foreach ($p in $points) {
    $x = $p.X * $scale
    if ($flip) { $x = -$x }
    [System.Drawing.PointF]::new($cx + $x, $cy + ($p.Y * $scale))
  }
  $g.FillPolygon($fill, $mapped)
  $g.DrawPolygon($outline, $mapped)
  $fingerX = 2
  if ($flip) { $fingerX = -2 }
  $g.DrawLine($accent, $cx + ($fingerX * $scale), $cy - 28 * $scale, $cx + ($fingerX * $scale), $cy - 10 * $scale)
}

function Draw-CycleSenseArtwork([System.Drawing.Graphics]$g, [int]$width, [int]$height, [bool]$forSplash) {
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.ColorTranslator]::FromHtml('#E7C1CB'))

  if ($forSplash) {
    $stroke = $width * 0.010
    $accentStroke = $width * 0.005
    $loopStroke = $width * 0.030
    $bodyWidth = $width * 0.24
    $bodyHeight = $height * 0.34
    $handScale = $width / 900.0
  } else {
    $stroke = $width * 0.024
    $accentStroke = $width * 0.012
    $loopStroke = $width * 0.075
    $bodyWidth = $width * 0.42
    $bodyHeight = $height * 0.40
    $handScale = $width / 430.0
  }

  $centerX = $width / 2.0
  $centerY = $height * 0.54

  $bodyBrush = New-Brush '#E85D8F'
  $bgBrush = New-Brush '#E7C1CB'
  $eyeBrush = New-Brush '#FFFFFF'
  $pupilBrush = New-Brush '#111111'
  $strokePen = New-Pen '#1E0E13' $stroke
  $accentPen = New-Pen '#D75134' $accentStroke
  $loopPen = New-Pen '#E85D8F' $loopStroke
  $loopAccentPen = New-Pen '#D75134' ($accentStroke * 1.6)

  $leftArc = [System.Drawing.RectangleF]::new($centerX - $bodyWidth * 1.55, $centerY - $bodyHeight * 0.95, $bodyWidth * 1.15, $bodyHeight * 1.05)
  $rightArc = [System.Drawing.RectangleF]::new($centerX + $bodyWidth * 0.40, $centerY - $bodyHeight * 0.95, $bodyWidth * 1.15, $bodyHeight * 1.05)
  $g.DrawArc($loopPen, $leftArc, 95, 315)
  $g.DrawArc($loopPen, $rightArc, -230, 315)
  $g.DrawArc($loopAccentPen, $leftArc.X + $accentStroke, $leftArc.Y + $accentStroke, $leftArc.Width - 2 * $accentStroke, $leftArc.Height - 2 * $accentStroke, 95, 315)
  $g.DrawArc($loopAccentPen, $rightArc.X + $accentStroke, $rightArc.Y + $accentStroke, $rightArc.Width - 2 * $accentStroke, $rightArc.Height - 2 * $accentStroke, -230, 315)

  $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
  $path.StartFigure()
  $path.AddBezier($centerX - $bodyWidth * 0.72, $centerY - $bodyHeight * 0.35, $centerX - $bodyWidth * 0.52, $centerY - $bodyHeight * 0.88, $centerX - $bodyWidth * 0.16, $centerY - $bodyHeight * 0.92, $centerX, $centerY - $bodyHeight * 0.62)
  $path.AddBezier($centerX, $centerY - $bodyHeight * 0.62, $centerX + $bodyWidth * 0.16, $centerY - $bodyHeight * 0.92, $centerX + $bodyWidth * 0.52, $centerY - $bodyHeight * 0.88, $centerX + $bodyWidth * 0.72, $centerY - $bodyHeight * 0.35)
  $path.AddBezier($centerX + $bodyWidth * 0.72, $centerY - $bodyHeight * 0.35, $centerX + $bodyWidth * 0.63, $centerY + $bodyHeight * 0.05, $centerX + $bodyWidth * 0.34, $centerY + $bodyHeight * 0.42, $centerX, $centerY + $bodyHeight * 0.96)
  $path.AddBezier($centerX, $centerY + $bodyHeight * 0.96, $centerX - $bodyWidth * 0.34, $centerY + $bodyHeight * 0.42, $centerX - $bodyWidth * 0.63, $centerY + $bodyHeight * 0.05, $centerX - $bodyWidth * 0.72, $centerY - $bodyHeight * 0.35)
  $path.CloseFigure()
  $g.FillPath($bodyBrush, $path)
  $g.DrawPath($strokePen, $path)

  $g.DrawArc($strokePen, $centerX - $bodyWidth * 0.19, $centerY - $bodyHeight * 0.80, $bodyWidth * 0.22, $bodyHeight * 0.17, 10, 150)
  $g.DrawArc($strokePen, $centerX + $bodyWidth * 0.02, $centerY - $bodyHeight * 0.78, $bodyWidth * 0.23, $bodyHeight * 0.19, 200, 140)

  $cervix = [System.Drawing.RectangleF]::new($centerX - $bodyWidth * 0.12, $centerY + $bodyHeight * 0.78, $bodyWidth * 0.24, $bodyHeight * 0.18)
  $g.FillEllipse($bodyBrush, $cervix)
  $g.DrawEllipse($strokePen, $cervix)
  $g.DrawArc($accentPen, $cervix.X + $cervix.Width * 0.1, $cervix.Y + $cervix.Height * 0.45, $cervix.Width * 0.8, $cervix.Height * 0.35, 10, 160)

  Draw-Hand $g ($centerX - $bodyWidth * 1.02) ($centerY - $bodyHeight * 0.03) $handScale $false $bgBrush $strokePen $accentPen
  Draw-Hand $g ($centerX + $bodyWidth * 1.02) ($centerY - $bodyHeight * 0.08) $handScale $true $bgBrush $strokePen $accentPen

  $leftEye = [System.Drawing.RectangleF]::new($centerX - $bodyWidth * 0.33, $centerY - $bodyHeight * 0.11, $bodyWidth * 0.19, $bodyWidth * 0.19)
  $g.FillEllipse($eyeBrush, $leftEye)
  $g.DrawEllipse($strokePen, $leftEye)
  $pupil = [System.Drawing.RectangleF]::new($leftEye.X + $leftEye.Width * 0.42, $leftEye.Y + $leftEye.Height * 0.42, $leftEye.Width * 0.16, $leftEye.Height * 0.16)
  $g.FillEllipse($pupilBrush, $pupil)
  $g.DrawArc($strokePen, $centerX + $bodyWidth * 0.02, $centerY - $bodyHeight * 0.01, $bodyWidth * 0.18, $bodyHeight * 0.11, 200, 140)
  $g.DrawArc($strokePen, $centerX + $bodyWidth * 0.01, $centerY - $bodyHeight * 0.10, $bodyWidth * 0.22, $bodyHeight * 0.14, 180, 160)
  $g.DrawArc($strokePen, $centerX - $bodyWidth * 0.06, $centerY + $bodyHeight * 0.10, $bodyWidth * 0.18, $bodyHeight * 0.12, 215, 110)
}

function Save-Art($path, $size, $splash) {
  $bmp = [System.Drawing.Bitmap]::new($size.Width, $size.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  Draw-CycleSenseArtwork $g $size.Width $size.Height $splash
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose()
  $bmp.Dispose()
}

$base = 'D:\Period Tracker\android\app\src\main\res'
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
  $size = New-Object System.Drawing.Size($entry.Value.Width, $entry.Value.Height)
  Save-Art (Join-Path $base $entry.Key) $size $false
}

$splashFiles = Get-ChildItem -Path $base -Recurse -Filter 'splash.png'
foreach ($file in $splashFiles) {
  $img = [System.Drawing.Image]::FromFile($file.FullName)
  $size = New-Object System.Drawing.Size($img.Width, $img.Height)
  $img.Dispose()
  Save-Art $file.FullName $size $true
}