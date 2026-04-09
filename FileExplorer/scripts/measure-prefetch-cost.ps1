$ws = New-Object -ComObject WScript.Shell
$sf = "C:\Users\szil\Shortcuts_to_folders"

# Resolve all shortcut targets
$targets = @()
foreach ($f in (Get-ChildItem $sf -ErrorAction SilentlyContinue)) {
    if ($f.PSIsContainer) {
        $targets += $f.FullName
    } elseif ($f.Extension -eq ".lnk") {
        try {
            $t = $ws.CreateShortcut($f.FullName).TargetPath
            if ($t -and (Test-Path $t -ErrorAction SilentlyContinue)) { $targets += $t }
        } catch {}
    }
}

Write-Host "Shortcut targets resolved: $($targets.Count)"
foreach ($t in $targets) { Write-Host "  $t" }

# Per-target 3-level measurement
$grandTotal = 0
$swAll = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($root in $targets) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $l1 = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)
    $dirs = $l1.Count
    $files = @(Get-ChildItem -LiteralPath $root -File -ErrorAction SilentlyContinue).Count
    foreach ($d1 in $l1) {
        $l2 = @(Get-ChildItem -LiteralPath $d1.FullName -Directory -ErrorAction SilentlyContinue)
        $dirs += $l2.Count
        foreach ($d2 in $l2) {
            $l3 = @(Get-ChildItem -LiteralPath $d2.FullName -Directory -ErrorAction SilentlyContinue)
            $dirs += $l3.Count
        }
    }
    $sw.Stop()
    $grandTotal += $dirs
    $shortName = Split-Path $root -Leaf
    Write-Host ("  {0,-40} dirs={1,4}  files_root={2,4}  ms={3}" -f $shortName,$dirs,$files,$sw.ElapsedMilliseconds)
}

$swAll.Stop()
Write-Host ""
Write-Host "TOTAL dirs across all targets (3 levels): $grandTotal"
Write-Host "TOTAL wall-clock time (serial): $($swAll.ElapsedMilliseconds) ms"
Write-Host "Estimated parallel time (4 threads): ~$([int]($swAll.ElapsedMilliseconds / 4)) ms"
