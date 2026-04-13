Add-Type -AssemblyName System.Drawing

$src  = "C:\Users\szil\Repos\excel_wizadry\FileExplorer\Explorer icon 256x256_.png"
$dest = "C:\Users\szil\Repos\excel_wizadry\FileExplorer\FileExplorer.App\Assets\LSApp.ico"

# Sizes: 256 stored as PNG-compressed (standard for large ICO), rest as BMP DIB (required for WM_SETICON / taskbar).
$sizes = @(256, 64, 48, 32, 24, 16)

function Get-IcoImageData([string]$pngPath, [int]$size) {
    $orig    = [System.Drawing.Bitmap]::new($pngPath)
    $resized = [System.Drawing.Bitmap]::new($orig, [System.Drawing.Size]::new($size, $size))
    $orig.Dispose()

    if ($size -ge 256) {
        # Store as PNG stream (smaller, widely supported for 256+)
        $ms = [System.IO.MemoryStream]::new()
        $resized.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $resized.Dispose()
        return $ms.ToArray()
    }

    # Store as BMP DIB — required for WM_SETICON / title bar / taskbar rendering.
    # ICO BMP format: BITMAPINFOHEADER + XOR mask (BGRA rows, bottom-up) + AND mask (1bpp, bottom-up).
    $w = $size; $h = $size
    $rowBytes   = $w * 4          # 32bpp BGRA
    $andRowBytes = [int]([Math]::Ceiling($w / 32.0)) * 4  # 1bpp, DWORD-aligned

    $dibMs = [System.IO.MemoryStream]::new()
    $bw    = [System.IO.BinaryWriter]::new($dibMs)

    # BITMAPINFOHEADER (40 bytes) — height is doubled for XOR+AND combo
    $bw.Write([int32]40)          # biSize
    $bw.Write([int32]$w)          # biWidth
    $bw.Write([int32]($h * 2))    # biHeight (XOR + AND stacked)
    $bw.Write([int16]1)           # biPlanes
    $bw.Write([int16]32)          # biBitCount
    $bw.Write([int32]0)           # biCompression = BI_RGB
    $bw.Write([int32]0)           # biSizeImage
    $bw.Write([int32]0)           # biXPelsPerMeter
    $bw.Write([int32]0)           # biYPelsPerMeter
    $bw.Write([int32]0)           # biClrUsed
    $bw.Write([int32]0)           # biClrImportant

    # XOR mask: 32bpp BGRA, bottom-up rows
    for ($row = $h - 1; $row -ge 0; $row--) {
        for ($col = 0; $col -lt $w; $col++) {
            $px = $resized.GetPixel($col, $row)
            $bw.Write([byte]$px.B)
            $bw.Write([byte]$px.G)
            $bw.Write([byte]$px.R)
            $bw.Write([byte]$px.A)
        }
    }

    # AND mask: 1bpp, bottom-up rows, pixel=0 means opaque (alpha from XOR), DWORD-aligned
    for ($row = $h - 1; $row -ge 0; $row--) {
        $andBytes = New-Object byte[] $andRowBytes
        # Leave all bits 0 (fully use alpha channel from XOR mask)
        $bw.Write($andBytes)
    }

    $bw.Flush()
    $resized.Dispose()
    return $dibMs.ToArray()
}

# Collect as List[byte[]] — prevents PowerShell from unrolling byte[] in foreach pipelines.
$imageData = [System.Collections.Generic.List[byte[]]]::new()
foreach ($sz in $sizes) { $imageData.Add((Get-IcoImageData $src $sz)) }

$ms  = [System.IO.MemoryStream]::new()
$bw  = [System.IO.BinaryWriter]::new($ms)

# ICO file header (6 bytes)
$bw.Write([int16]0)                   # reserved
$bw.Write([int16]1)                   # type = icon
$bw.Write([int16]$sizes.Count)

# Directory entries (16 bytes each)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $sz  = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
    $bw.Write([byte]$sz)
    $bw.Write([byte]$sz)
    $bw.Write([byte]0)   # colour count
    $bw.Write([byte]0)   # reserved
    $bw.Write([int16]1)  # planes
    $bw.Write([int16]32) # bit depth
    $bw.Write([int32]$imageData[$i].Length)
    $bw.Write([int32]$offset)
    $offset += $imageData[$i].Length
}

foreach ($d in $imageData) { $bw.Write($d) }

$bw.Flush()
[System.IO.File]::WriteAllBytes($dest, $ms.ToArray())
Write-Host "Done: $dest  ($([System.IO.File]::ReadAllBytes($dest).Length) bytes)"
