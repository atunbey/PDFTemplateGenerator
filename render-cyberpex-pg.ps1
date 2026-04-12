$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$basePath = 'PDFTemplateGenerator\wwwroot\images\cyberpex-banner-section-bg.png'
$outPath = 'PDFTemplateGenerator\wwwroot\images\cyberpex-banner-section-pg.png'
if (-not (Test-Path $basePath)) { throw "Base image not found: $basePath" }

$bmp = [System.Drawing.Bitmap]::FromFile((Resolve-Path $basePath))
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

# 1) Upper sky section
$rectTop = New-Object System.Drawing.Rectangle 0,0,$bmp.Width,420
$brushTop = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rectTop,[System.Drawing.Color]::FromArgb(58,5,2,10),[System.Drawing.Color]::FromArgb(0,18,7,31),90)
$g.FillRectangle($brushTop,$rectTop)
$brushTop.Dispose()

# 2) Mid sky stars with fade by distance from planet
$rand = New-Object System.Random 1337
for ($i=0; $i -lt 180; $i++) {
  $x = $rand.Next(20,$bmp.Width-20)
  $y = $rand.Next(80,560)
  $alpha = [Math]::Min(120,[int](8 + (($y-80)/480.0)*112))
  $r = if ($rand.NextDouble() -lt 0.12) { 2.0 } elseif ($rand.NextDouble() -lt 0.45) { 1.5 } else { 1.0 }
  $glowA = [Math]::Max(6,[int]($alpha*0.30))
  $glowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($glowA,245,185,255))
  $coreBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($alpha,246,200,255))
  $g.FillEllipse($glowBrush,[float]($x-($r*1.6)),[float]($y-($r*1.6)),[float]($r*3.2),[float]($r*3.2))
  $g.FillEllipse($coreBrush,[float]($x-$r),[float]($y-$r),[float]($r*2),[float]($r*2))
  $glowBrush.Dispose()
  $coreBrush.Dispose()
}

# 3) Sirius A and B in middle sky
$sx = 960
$sy = 336
$saGlow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(90,255,230,255))
$saCore = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(210,255,242,255))
$sbGlow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(62,255,210,255))
$sbCore = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(185,255,228,255))
$g.FillEllipse($saGlow,$sx-11,$sy-11,22,22)
$g.FillEllipse($saCore,$sx-4,$sy-4,8,8)
$g.FillEllipse($sbGlow,$sx+15,$sy+2,14,14)
$g.FillEllipse($sbCore,$sx+20,$sy+7,5,5)
$starPenA = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(148,255,236,255),1.2)
$starPenB = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(118,255,224,255),1.0)
$g.DrawLine($starPenA,$sx-16,$sy,$sx+16,$sy)
$g.DrawLine($starPenA,$sx,$sy-16,$sx,$sy+16)
$g.DrawLine($starPenB,$sx+10,$sy+9,$sx+35,$sy+9)
$g.DrawLine($starPenB,$sx+22,$sy-3,$sx+22,$sy+21)
$saGlow.Dispose(); $saCore.Dispose(); $sbGlow.Dispose(); $sbCore.Dispose(); $starPenA.Dispose(); $starPenB.Dispose()

# 4) Sunrise section
$sunRect = New-Object System.Drawing.Rectangle 520,540,880,320
$path = New-Object System.Drawing.Drawing2D.GraphicsPath
$path.AddEllipse($sunRect)
$pgb = New-Object System.Drawing.Drawing2D.PathGradientBrush($path)
$pgb.CenterColor = [System.Drawing.Color]::FromArgb(86,255,162,246)
$pgb.SurroundColors = @([System.Drawing.Color]::FromArgb(0,186,70,255))
$g.FillPath($pgb,$path)
$horizonPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(86,210,92,255),3.0)
$g.DrawArc($horizonPen,260,636,1400,280,196,148)
$path.Dispose(); $pgb.Dispose(); $horizonPen.Dispose()

# 5) Planet cap cue with Africa centered and visible
$africaGlow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(42,227,126,255))
$g.FillEllipse($africaGlow,800,710,320,150)
$africaPoints = @(
  (New-Object System.Drawing.Point 938,710), (New-Object System.Drawing.Point 916,718), (New-Object System.Drawing.Point 894,744),
  (New-Object System.Drawing.Point 899,774), (New-Object System.Drawing.Point 914,801), (New-Object System.Drawing.Point 921,832),
  (New-Object System.Drawing.Point 934,860), (New-Object System.Drawing.Point 952,888), (New-Object System.Drawing.Point 974,913),
  (New-Object System.Drawing.Point 1003,937), (New-Object System.Drawing.Point 1024,957), (New-Object System.Drawing.Point 1042,981),
  (New-Object System.Drawing.Point 1056,968), (New-Object System.Drawing.Point 1064,941), (New-Object System.Drawing.Point 1080,927),
  (New-Object System.Drawing.Point 1096,903), (New-Object System.Drawing.Point 1107,871), (New-Object System.Drawing.Point 1124,846),
  (New-Object System.Drawing.Point 1144,832), (New-Object System.Drawing.Point 1140,806), (New-Object System.Drawing.Point 1119,790),
  (New-Object System.Drawing.Point 1092,778), (New-Object System.Drawing.Point 1068,756), (New-Object System.Drawing.Point 1053,724),
  (New-Object System.Drawing.Point 1042,690), (New-Object System.Drawing.Point 1014,667), (New-Object System.Drawing.Point 983,659),
  (New-Object System.Drawing.Point 956,664)
)
$africaBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(52,223,132,255))
$g.FillPolygon($africaBrush,$africaPoints)
$africaGlow.Dispose(); $africaBrush.Dispose()

$g.Dispose()
$bmp.Save((Resolve-Path 'PDFTemplateGenerator\wwwroot\images').Path + '\cyberpex-banner-section-pg.png',[System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

Get-Item $outPath | Select-Object FullName, Length
