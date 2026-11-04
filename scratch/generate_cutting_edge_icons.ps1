Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = (Resolve-Path (Join-Path $scriptDir "..")).Path
$srcIconPath = Join-Path $rootDir "expenses_14510803.png"

if (-not (Test-Path $srcIconPath)) {
    Write-Error "Source icon not found at: $srcIconPath"
    exit 1
}

Write-Host "================================================================================"
Write-Host "GFMP CUTTING-EDGE ICON GENERATOR - PRODUCTION v1.01"
Write-Host "Source: expenses_14510803.png (512x512 Master Asset)"
Write-Host "================================================================================"

$srcBmp = [System.Drawing.Bitmap]::FromFile($srcIconPath)

function Resize-Icon([System.Drawing.Bitmap]$source, [int]$w, [int]$h, [string]$outputPath) {
    $dir = Split-Path -Parent $outputPath
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $destBmp = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($destBmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
    $g.DrawImage($source, $rect, 0, 0, $source.Width, $source.Height, [System.Drawing.GraphicsUnit]::Pixel)
    $g.Dispose()

    $destBmp.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $destBmp.Dispose()
    Write-Host "  -> Generated: $outputPath (${w}x${h})"
}

# 1. Update master icon at root
$masterPath = Join-Path $rootDir "master_gfmp_icon.png"
Copy-Item $srcIconPath $masterPath -Force
Write-Host "`n1. Master Icon at root updated: master_gfmp_icon.png"

# 2. Update Web/PWA Assets
Write-Host "`n2. Generating Web/PWA Assets (mobile-web/assets)..."
$mobileWebAssets = Join-Path $rootDir "mobile-web\assets"

Resize-Icon $srcBmp 512 512 (Join-Path $mobileWebAssets "icon_512.png")
Resize-Icon $srcBmp 512 512 (Join-Path $mobileWebAssets "logo_jic.png")
Resize-Icon $srcBmp 512 512 (Join-Path $mobileWebAssets "logo_jic_transparent.png")
Resize-Icon $srcBmp 192 192 (Join-Path $mobileWebAssets "favicon_jic.png")

# 3. Generate Android Mipmap Icons
Write-Host "`n3. Generating Android Mipmap Icons across all densities..."
$resDir = Join-Path $rootDir "android\app\src\main\res"

$densities = @(
    @{ Name = "mipmap-mdpi"; Size = 48 },
    @{ Name = "mipmap-hdpi"; Size = 72 },
    @{ Name = "mipmap-xhdpi"; Size = 96 },
    @{ Name = "mipmap-xxhdpi"; Size = 144 },
    @{ Name = "mipmap-xxxhdpi"; Size = 192 }
)

foreach ($d in $densities) {
    $folder = Join-Path $resDir $d.Name
    $icLauncher = Join-Path $folder "ic_launcher.png"
    $icLauncherRound = Join-Path $folder "ic_launcher_round.png"
    Resize-Icon $srcBmp $d.Size $d.Size $icLauncher
    Resize-Icon $srcBmp $d.Size $d.Size $icLauncherRound
}

$srcBmp.Dispose()

Write-Host "`n================================================================================"
Write-Host "ALL CUTTING-EDGE ICONS SUCCESSFULLY GENERATED!"
Write-Host "================================================================================"
