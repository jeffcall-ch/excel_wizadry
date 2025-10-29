# Icon Generation Script

Since binary .ico files cannot be created directly through text, you have several options:

## Option 1: Use an Existing Icon
Copy any .ico file you like and rename it to `appicon.ico` in this directory.

## Option 2: Convert PNG to ICO Online
1. Create or find a 256x256 PNG image
2. Visit https://convertio.co/png-ico/ or https://www.icoconverter.com/
3. Upload your PNG and convert to ICO
4. Save as `appicon.ico` in this directory

## Option 3: Use Windows Built-in Icons
You can extract icons from Windows system files:

```powershell
# Extract an icon from a system file (requires additional tools)
# Or simply copy from C:\Windows\System32\*.ico files
```

## Option 4: Generate Programmatically
Use the PowerShell script below to create a simple colored icon:

```powershell
# Create a simple icon using .NET
Add-Type -AssemblyName System.Drawing

$size = 256
$bitmap = New-Object System.Drawing.Bitmap($size, $size)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# Fill with a color
$brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::DodgerBlue)
$graphics.FillRectangle($brush, 0, 0, $size, $size)

# Draw a folder shape
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 8)
$graphics.DrawRectangle($pen, 40, 80, 176, 140)
$graphics.DrawRectangle($pen, 40, 60, 80, 20)

$graphics.Dispose()

# Convert to icon (simplified - actual icon creation requires more steps)
# Save as PNG first, then convert to ICO using online tool
$bitmap.Save("$PSScriptRoot\appicon.png", [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()

Write-Host "PNG created. Please convert appicon.png to appicon.ico using an online converter."
```

## Option 5: Use Default Windows App Icon (Temporary)
The app will use a default icon if `appicon.ico` is missing, but it's recommended to add a custom one.

## Recommended Tool
**IconsFlow** or **Greenfish Icon Editor Pro** (free) - allows creating multi-resolution .ico files.

---

**Note:** For now, you can build and test the app without a custom icon. Add `appicon.ico` before publishing for distribution.
