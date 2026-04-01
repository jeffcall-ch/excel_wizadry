param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",

    [ValidateSet("win-x64")]
    [string]$RuntimeIdentifier = "win-x64",

    [switch]$SkipZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot "FileExplorer.App\FileExplorer.App.csproj"
$publishDir = Join-Path $repoRoot "artifacts\portable-folder\$RuntimeIdentifier\FileExplorer"
$zipPath = Join-Path $repoRoot "artifacts\portable-folder\$RuntimeIdentifier\FileExplorer-portable-$RuntimeIdentifier.zip"
$runtimeBuildDir = Join-Path $repoRoot "FileExplorer.App\bin\$Configuration\net8.0-windows10.0.22621.0\$RuntimeIdentifier"

Get-Process FileExplorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

if (Test-Path $publishDir) {
    Remove-Item -Path $publishDir -Recurse -Force
}

New-Item -Path $publishDir -ItemType Directory -Force | Out-Null

$publishArgs = @(
    "publish",
    $projectPath,
    "-c", $Configuration,
    "-r", $RuntimeIdentifier,
    "--self-contained", "true",
    "/p:PublishProfile=Properties\PublishProfiles\PortableFolder-x64.pubxml",
    "/p:PortableFolderPublish=true",
    "/p:PublishDir=$publishDir\"
)

Write-Host "Publishing portable folder..."
Write-Host "dotnet $($publishArgs -join ' ')"

& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $runtimeBuildDir)) {
    throw "Expected runtime build output was not found: $runtimeBuildDir"
}

# Publish layout has shown startup regressions for this WinUI app.
# Assemble the portable package from the known-good runtime build output.
Remove-Item -Path $publishDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -Path $publishDir -ItemType Directory -Force | Out-Null
Copy-Item -Path (Join-Path $runtimeBuildDir "*") -Destination $publishDir -Recurse -Force

$requiredFiles = @(
    "FileExplorer.exe",
    "resources.pri",
    "MainWindow.xaml",
    "Views\FileListView.xaml"
)

$missingFiles = @()
foreach ($file in $requiredFiles) {
    if (-not (Test-Path (Join-Path $publishDir $file))) {
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    throw "Portable publish is missing required file(s): $($missingFiles -join ', ')"
}

if (-not $SkipZip) {
    if (Test-Path $zipPath) {
        Remove-Item -Path $zipPath -Force
    }

    Compress-Archive -Path (Join-Path $publishDir "*") -DestinationPath $zipPath -CompressionLevel Optimal
}

Write-Host ""
Write-Host "Portable folder publish completed successfully."
Write-Host "Folder: $publishDir"
if (-not $SkipZip) {
    Write-Host "Zip:    $zipPath"
}
