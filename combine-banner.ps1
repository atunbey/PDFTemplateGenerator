$ErrorActionPreference = 'Stop'

$root      = Split-Path -Parent $MyInvocation.MyCommand.Path
$imgDir    = Join-Path $root 'PDFTemplateGenerator\wwwroot\images'
$bgPath    = Join-Path $imgDir 'cyberpex-banner-section-bg.png'
$svgPath   = Join-Path $imgDir 'weighing-of-the-heart.svg'
$outPath   = Join-Path $imgDir 'cyberpex-banner-combined.png'
$tmpDir    = Join-Path $root '_tmp_combine'

if (-not (Test-Path $bgPath))  { throw "Background not found: $bgPath" }
if (-not (Test-Path $svgPath)) { throw "SVG not found: $svgPath" }

# ---------- create temp project ----------
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

$csproj = @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net9.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="SkiaSharp"           Version="2.88.8" />
    <PackageReference Include="SkiaSharp.NativeAssets.Win32" Version="2.88.8" />
    <PackageReference Include="Svg.Skia"            Version="1.0.0.3" />
  </ItemGroup>
</Project>
'@

$program = @'
using SkiaSharp;
using Svg.Skia;

var args_ = args;
string bgPath  = args_[0];
string svgPath = args_[1];
string outPath = args_[2];

// --- load background PNG ---
using var bgBmp = SKBitmap.Decode(bgPath)
    ?? throw new Exception("Failed to decode background PNG");

int W = bgBmp.Width;   // 1920
int H = bgBmp.Height;  // 1188

using var surface = SKSurface.Create(new SKImageInfo(W, H, SKColorType.Rgba8888, SKAlphaType.Premul));
var canvas = surface.Canvas;
canvas.DrawBitmap(bgBmp, 0, 0);

// --- render SVG on top, centred, 1400 px wide ---
var svg = new SKSvg();
svg.Load(svgPath);

var bounds = svg.Picture!.CullRect;
float targetW = 1400f;
float scale   = targetW / bounds.Width;
float targetH = bounds.Height * scale;

float dx = (W - targetW) / 2f;
float dy = (H - targetH) / 2f;

var matrix = SKMatrix.CreateScaleTranslation(scale, scale, dx, dy);
canvas.DrawPicture(svg.Picture, ref matrix);

// --- save ---
using var img  = surface.Snapshot();
using var data = img.Encode(SKEncodedImageFormat.Png, 92);
using var fs   = File.OpenWrite(outPath);
data.SaveTo(fs);
Console.WriteLine("Saved: " + outPath);
'@

Set-Content -Path (Join-Path $tmpDir 'Combine.csproj') -Value $csproj -Encoding UTF8
Set-Content -Path (Join-Path $tmpDir 'Program.cs')     -Value $program  -Encoding UTF8

# ---------- run ----------
Write-Host 'Restoring packages...'
& dotnet restore (Join-Path $tmpDir 'Combine.csproj') --verbosity quiet

Write-Host 'Building and running...'
& dotnet run --project (Join-Path $tmpDir 'Combine.csproj') -c Release -- $bgPath $svgPath $outPath

# ---------- clean up ----------
Remove-Item -Recurse -Force $tmpDir

if (Test-Path $outPath) {
    $info = [System.IO.FileInfo]$outPath
    Write-Host ("Done! Output: {0}  ({1:N0} KB)" -f $outPath, ($info.Length / 1KB))
} else {
    throw "Output file was not created."
}
