Add-Type -AssemblyName System.Drawing

function New-LuxuryMedallionBitmap([int]$size) {
    $ss = $size * 2
    $bmpSS = New-Object System.Drawing.Bitmap($ss, $ss, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmpSS)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    $scale = $ss / 64.0

    # 1. Outer Gold Medallion Rim
    $rectOuter = New-Object System.Drawing.RectangleF(([float](2.0 * $scale)), ([float](2.0 * $scale)), ([float](60.0 * $scale)), ([float](60.0 * $scale)))
    $brushGoldRim = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF(0, 0)),
        (New-Object System.Drawing.PointF($ss, $ss)),
        [System.Drawing.Color]::FromArgb(255, 255, 246, 183),
        [System.Drawing.Color]::FromArgb(255, 165, 118, 20)
    )
    $cbRim = New-Object System.Drawing.Drawing2D.ColorBlend(4)
    $cbRim.Colors = @(
        [System.Drawing.Color]::FromArgb(255, 255, 249, 204),
        [System.Drawing.Color]::FromArgb(255, 250, 218, 94),
        [System.Drawing.Color]::FromArgb(255, 212, 175, 55),
        [System.Drawing.Color]::FromArgb(255, 150, 105, 18)
    )
    $cbRim.Positions = @(0.0, 0.35, 0.72, 1.0)
    $brushGoldRim.InterpolationColors = $cbRim
    $g.FillEllipse($brushGoldRim, $rectOuter)

    # 2. Inner Royal Sapphire Medallion Core
    $rectCore = New-Object System.Drawing.RectangleF(([float](4.5 * $scale)), ([float](4.5 * $scale)), ([float](55.0 * $scale)), ([float](55.0 * $scale)))
    $brushSapphire = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF(0, 0)),
        (New-Object System.Drawing.PointF($ss, $ss)),
        [System.Drawing.Color]::FromArgb(255, 30, 62, 150),
        [System.Drawing.Color]::FromArgb(255, 6, 12, 32)
    )
    $cbSap = New-Object System.Drawing.Drawing2D.ColorBlend(3)
    $cbSap.Colors = @(
        [System.Drawing.Color]::FromArgb(255, 34, 68, 162),
        [System.Drawing.Color]::FromArgb(255, 14, 28, 68),
        [System.Drawing.Color]::FromArgb(255, 6, 12, 32)
    )
    $cbSap.Positions = @(0.0, 0.58, 1.0)
    $brushSapphire.InterpolationColors = $cbSap
    $g.FillEllipse($brushSapphire, $rectCore)

    # 3. Inner Fine Gold Ring (R = 24.4 -> xy = 7.6 .. 56.4)
    $penInnerRing = New-Object System.Drawing.Pen($brushGoldRim, ([float](0.85 * $scale)))
    $rectInnerRing = New-Object System.Drawing.RectangleF(([float](7.6 * $scale)), ([float](7.6 * $scale)), ([float](48.8 * $scale)), ([float](48.8 * $scale)))
    $g.DrawEllipse($penInnerRing, $rectInnerRing)

    # 4. Cardinal Diamond Studs on the Rim Band
    $brushBrightGold = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 248, 200))
    $studCoords = @(
        @(32.0, 6.0),
        @(32.0, 58.0),
        @(6.0, 32.0),
        @(58.0, 32.0)
    )
    foreach ($sc in $studCoords) {
        $rStud = 1.2 * $scale
        $g.FillEllipse($brushBrightGold, ([float]($sc[0] * $scale - $rStud)), ([float]($sc[1] * $scale - $rStud)), ([float]($rStud * 2)), ([float]($rStud * 2)))
    }

    # 5. Sovereign Cut Diamond Crest at Top Center (Crown cut: top flat table + side facets + bottom pavilion point)
    $diamondPoly = [System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF(([float](28.2 * $scale)), ([float](12.2 * $scale)))),
        (New-Object System.Drawing.PointF(([float](35.8 * $scale)), ([float](12.2 * $scale)))),
        (New-Object System.Drawing.PointF(([float](38.8 * $scale)), ([float](15.8 * $scale)))),
        (New-Object System.Drawing.PointF(([float](32.0 * $scale)), ([float](23.5 * $scale)))),
        (New-Object System.Drawing.PointF(([float](25.2 * $scale)), ([float](15.8 * $scale))))
    )
    $g.FillPolygon($brushGoldRim, $diamondPoly)
    $penFacet = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(240, 255, 250, 220), ([float](0.8 * $scale)))
    $g.DrawPolygon($penFacet, $diamondPoly)
    # Girdle line & internal crown/pavilion facets
    $g.DrawLine($penFacet, ([float](25.4 * $scale)), ([float](15.8 * $scale)), ([float](38.6 * $scale)), ([float](15.8 * $scale)))
    $g.DrawLine($penFacet, ([float](28.2 * $scale)), ([float](12.2 * $scale)), ([float](29.4 * $scale)), ([float](15.8 * $scale)))
    $g.DrawLine($penFacet, ([float](35.8 * $scale)), ([float](12.2 * $scale)), ([float](34.6 * $scale)), ([float](15.8 * $scale)))
    $g.DrawLine($penFacet, ([float](29.4 * $scale)), ([float](15.8 * $scale)), ([float](32.0 * $scale)), ([float](23.3 * $scale)))
    $g.DrawLine($penFacet, ([float](34.6 * $scale)), ([float](15.8 * $scale)), ([float](32.0 * $scale)), ([float](23.3 * $scale)))

    # 6. Symmetrical Ascending Private Banking Wealth Pillars (Cleanly inside inner ring: x = 18.5..45.5, y = 26.5..44.5)
    $bars = @(
        @(18.8, 35.5, 6.4, 9.0),
        @(28.8, 31.0, 6.4, 13.5),
        @(38.8, 25.5, 6.4, 19.0)
    )
    foreach ($b in $bars) {
        $bx = [float]($b[0] * $scale)
        $by = [float]($b[1] * $scale)
        $bw = [float]($b[2] * $scale)
        $bh = [float]($b[3] * $scale)
        $g.FillRectangle($brushGoldRim, $bx, $by, $bw, $bh)
        $g.DrawRectangle($penFacet, $bx, $by, $bw, $bh)
    }

    # 7. Pedestal Base Lines (Strictly inside the inner circle: x = 16.5..47.5, y = 46.8 and x = 21.0..43.0, y = 49.4)
    $penBase1 = New-Object System.Drawing.Pen($brushGoldRim, ([float](2.0 * $scale)))
    $penBase1.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $penBase1.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($penBase1, ([float](16.5 * $scale)), ([float](46.8 * $scale)), ([float](47.5 * $scale)), ([float](46.8 * $scale)))

    $penBase2 = New-Object System.Drawing.Pen($brushGoldRim, ([float](1.25 * $scale)))
    $penBase2.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $penBase2.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($penBase2, ([float](21.5 * $scale)), ([float](49.5 * $scale)), ([float](42.5 * $scale)), ([float](49.5 * $scale)))

    $g.Dispose()

    $bmpOut = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $gOut = [System.Drawing.Graphics]::FromImage($bmpOut)
    $gOut.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $gOut.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $gOut.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $gOut.DrawImage($bmpSS, 0, 0, $size, $size)
    $gOut.Dispose()
    $bmpSS.Dispose()

    return $bmpOut
}

function Save-IconPng([int]$size, [string]$outPath) {
    $bmp = New-LuxuryMedallionBitmap $size
    $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "Saved $size x $size -> $outPath"
}

$repoRoot = "c:\Users\KIMYAÉDITION\source\repos\GFMP"

Save-IconPng 192 "$repoRoot\mobile-web\assets\favicon_jic.png"
Save-IconPng 512 "$repoRoot\mobile-web\assets\icon_512.png"
Save-IconPng 512 "$repoRoot\mobile-web\assets\logo_jic.png"
Save-IconPng 512 "$repoRoot\mobile-web\assets\logo_jic_transparent.png"

$mipmaps = @{
    "mipmap-mdpi"    = 48
    "mipmap-hdpi"    = 72
    "mipmap-xhdpi"   = 96
    "mipmap-xxhdpi"  = 144
    "mipmap-xxxhdpi" = 192
}
foreach ($kv in $mipmaps.GetEnumerator()) {
    $dir = "$repoRoot\android\app\src\main\res\$($kv.Key)"
    Save-IconPng $kv.Value "$dir\ic_launcher.png"
    Save-IconPng $kv.Value "$dir\ic_launcher_round.png"
}
