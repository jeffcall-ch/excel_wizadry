param(
    [string]$CacheDirectory = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($CacheDirectory)) {
    $CacheDirectory = Join-Path $repoRoot "Development"
}

$dbPath = Join-Path $CacheDirectory "identity_cache.db"
$shmPath = "$dbPath-shm"
$walPath = "$dbPath-wal"

Write-Host "Resetting SQLite cache files..."
Write-Host "Directory: $CacheDirectory"

Get-Process FileExplorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$paths = @($dbPath, $shmPath, $walPath)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
        if (Test-Path $path) {
            Write-Warning "Failed to delete: $path"
        }
        else {
            Write-Host "Deleted: $path"
        }
    }
    else {
        Write-Host "Missing (ok): $path"
    }
}

Write-Host "Cache reset completed."
