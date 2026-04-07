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

function Reset-DirectoryContents {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        return
    }

    try {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        return
    }
    catch {
        # Keep the root directory and clear child items if the root is locked (for example, as a current working directory).
    }

    Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue |
        ForEach-Object {
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
}

Reset-DirectoryContents -Path $publishDir

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
Reset-DirectoryContents -Path $publishDir
Copy-Item -Path (Join-Path $runtimeBuildDir "*") -Destination $publishDir -Recurse -Force

$allowedProjectsCsv = Join-Path $repoRoot "Development\allowed_project_names.csv"
if (Test-Path $allowedProjectsCsv) {
    $publishDevDir = Join-Path $publishDir "Development"
    if (-not (Test-Path $publishDevDir)) {
        New-Item -Path $publishDevDir -ItemType Directory -Force | Out-Null
    }

    Copy-Item -Path $allowedProjectsCsv -Destination (Join-Path $publishDevDir "allowed_project_names.csv") -Force
    Copy-Item -Path $allowedProjectsCsv -Destination (Join-Path $publishDir "allowed_project_names.csv") -Force
}

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
