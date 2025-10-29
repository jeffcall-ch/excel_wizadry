# Generate a basic placeholder icon for the Shortcut Launcher app
# Run this script in PowerShell to create appicon.png
# Then convert PNG to ICO using an online converter

param(
    [string]$OutputPath = "$PSScriptRoot\appicon.png"
)

Write-Host "Generating placeholder icon..." -ForegroundColor Cyan

try {
    Add-Type -AssemblyName System.Drawing

    $size = 256
    $bitmap = New-Object System.Drawing.Bitmap($size, $size)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    
    # Enable anti-aliasing
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    
    # Background - gradient blue
    $gradientBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(0, 0)),
        (New-Object System.Drawing.Point($size, $size)),
        [System.Drawing.Color]::FromArgb(33, 150, 243),  # Light blue
        [System.Drawing.Color]::FromArgb(21, 101, 192)   # Darker blue
    )
    $graphics.FillRectangle($gradientBrush, 0, 0, $size, $size)
    
    # Draw a folder icon
    $folderColor = [System.Drawing.Color]::FromArgb(255, 235, 59)  # Yellow
    $folderBrush = New-Object System.Drawing.SolidBrush($folderColor)
    
    # Folder back
    $graphics.FillRectangle($folderBrush, 48, 96, 160, 120)
    
    # Folder tab
    $graphics.FillRectangle($folderBrush, 48, 72, 80, 24)
    
    # Add shadow/outline
    $outlinePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 160, 40), 3)
    $graphics.DrawRectangle($outlinePen, 48, 96, 160, 120)
    $graphics.DrawRectangle($outlinePen, 48, 72, 80, 24)
    
    # Draw shortcut arrow
    $arrowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $arrowPoints = @(
        (New-Object System.Drawing.Point(140, 140)),
        (New-Object System.Drawing.Point(180, 140)),
        (New-Object System.Drawing.Point(180, 120)),
        (New-Object System.Drawing.Point(200, 145)),
        (New-Object System.Drawing.Point(180, 170)),
        (New-Object System.Drawing.Point(180, 150)),
        (New-Object System.Drawing.Point(140, 150))
    )
    $graphics.FillPolygon($arrowBrush, $arrowPoints)
    
    # Add border outline
    $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(100, 255, 255, 255), 4)
    $graphics.DrawRectangle($borderPen, 2, 2, ($size - 4), ($size - 4))
    
    $graphics.Dispose()
    $gradientBrush.Dispose()
    $folderBrush.Dispose()
    $arrowBrush.Dispose()
    $outlinePen.Dispose()
    $borderPen.Dispose()
    
    # Save PNG
    $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
    
    Write-Host "✅ Icon PNG created successfully: $OutputPath" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Convert this PNG to ICO format using an online tool:" -ForegroundColor White
    Write-Host "   - https://convertio.co/png-ico/" -ForegroundColor Cyan
    Write-Host "   - https://www.icoconverter.com/" -ForegroundColor Cyan
    Write-Host "2. Save the ICO file as 'appicon.ico' in this directory" -ForegroundColor White
    Write-Host "3. Rebuild the project" -ForegroundColor White
    
    # Try to open the PNG
    if (Test-Path $OutputPath) {
        Start-Process $OutputPath
    }
}
catch {
    Write-Host "❌ Error generating icon: $_" -ForegroundColor Red
    Write-Host "You can manually create an icon or use a placeholder." -ForegroundColor Yellow
}
