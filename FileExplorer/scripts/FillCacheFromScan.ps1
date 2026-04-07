param(
    [string]$ShortcutsFolder = "",
    [string]$EnginePath = "",
    [string]$CacheDirectory = "",
    [int]$TimeoutSeconds = 1800
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

function Resolve-ShortcutsFolder {
    param([string]$Candidate)

    if (-not [string]::IsNullOrWhiteSpace($Candidate) -and (Test-Path $Candidate)) {
        return (Resolve-Path $Candidate).Path
    }

    $envPath = $env:FILEEXPLORER_SHORTCUTS_FOLDER
    if (-not [string]::IsNullOrWhiteSpace($envPath)) {
        $expanded = [Environment]::ExpandEnvironmentVariables($envPath)
        if (Test-Path $expanded) {
            return (Resolve-Path $expanded).Path
        }
    }

    $userDefault = Join-Path $env:USERPROFILE "Shortcuts_to_folders"
    if (Test-Path $userDefault) {
        return (Resolve-Path $userDefault).Path
    }

    $nearbyPlural = Join-Path $repoRoot "Shortcuts_to_folders"
    if (Test-Path $nearbyPlural) {
        return (Resolve-Path $nearbyPlural).Path
    }

    $nearbySingular = Join-Path $repoRoot "Shortcut_to_folders"
    if (Test-Path $nearbySingular) {
        return (Resolve-Path $nearbySingular).Path
    }

    throw "No shortcuts folder found. Set -ShortcutsFolder or FILEEXPLORER_SHORTCUTS_FOLDER."
}

function Resolve-EnginePath {
    param([string]$Candidate)

    if (-not [string]::IsNullOrWhiteSpace($Candidate) -and (Test-Path $Candidate)) {
        return (Resolve-Path $Candidate).Path
    }

    if (-not [string]::IsNullOrWhiteSpace($env:FILEEXPLORER_RQ_SEARCH_PATH) -and (Test-Path $env:FILEEXPLORER_RQ_SEARCH_PATH)) {
        return (Resolve-Path $env:FILEEXPLORER_RQ_SEARCH_PATH).Path
    }

    $candidates = @(
        (Join-Path $repoRoot "FileExplorer.App\Assets\rq_search.exe"),
        (Join-Path $repoRoot "FileExplorer.App\Assets\rq.exe"),
        (Join-Path $repoRoot "artifacts\portable-folder\win-x64\FileExplorer\Assets\rq_search.exe"),
        (Join-Path $repoRoot "artifacts\portable-folder\win-x64\FileExplorer\Assets\rq.exe")
    )

    foreach ($item in $candidates) {
        if (Test-Path $item) {
            return (Resolve-Path $item).Path
        }
    }

    throw "No rq engine found. Set -EnginePath or FILEEXPLORER_RQ_SEARCH_PATH."
}

function Get-FavoriteRoots {
    param([string]$Folder)

    $roots = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $shell = $null

    try {
        $shell = New-Object -ComObject WScript.Shell
    }
    catch {
        $shell = $null
    }

    Get-ChildItem -Path $Folder -Force -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.PSIsContainer) {
            [void]$roots.Add($_.FullName)
            return
        }

        if ($_.Extension -ieq ".lnk" -and $shell -ne $null) {
            try {
                $shortcut = $shell.CreateShortcut($_.FullName)
                $target = $shortcut.TargetPath
                if (-not [string]::IsNullOrWhiteSpace($target) -and (Test-Path $target) -and (Get-Item $target).PSIsContainer) {
                    [void]$roots.Add((Resolve-Path $target).Path)
                }
            }
            catch {
                # Ignore malformed or inaccessible shortcut entries.
            }
        }
    }

    return @($roots)
}

$resolvedShortcutsFolder = Resolve-ShortcutsFolder -Candidate $ShortcutsFolder
$resolvedEnginePath = Resolve-EnginePath -Candidate $EnginePath
$favoritesRoots = Get-FavoriteRoots -Folder $resolvedShortcutsFolder

if ([string]::IsNullOrWhiteSpace($CacheDirectory)) {
    $CacheDirectory = Join-Path $repoRoot "Development"
}

$resolvedCacheDirectory = ""
$resolvedPath = Resolve-Path -Path $CacheDirectory -ErrorAction SilentlyContinue
if ($resolvedPath -ne $null) {
    $resolvedCacheDirectory = $resolvedPath.Path
}

if ([string]::IsNullOrWhiteSpace($resolvedCacheDirectory)) {
    New-Item -Path $CacheDirectory -ItemType Directory -Force | Out-Null
    $resolvedCacheDirectory = (Resolve-Path -Path $CacheDirectory).Path
}

$resolvedCacheDbPath = Join-Path $resolvedCacheDirectory "identity_cache.db"

if ($favoritesRoots.Count -eq 0) {
    throw "No favorite roots discovered in $resolvedShortcutsFolder"
}

$env:FILEEXPLORER_RQ_SEARCH_PATH = $resolvedEnginePath
$env:FILEEXPLORER_IDENTITY_CACHE_DB = $resolvedCacheDbPath

Write-Host "Filling identity cache from deep scan..."
Write-Host "Shortcuts folder: $resolvedShortcutsFolder"
Write-Host "Engine path: $resolvedEnginePath"
Write-Host "Cache DB path: $resolvedCacheDbPath"
Write-Host "Favorite roots: $($favoritesRoots.Count)"

$tmp = Join-Path $env:TEMP ("cache_fill_" + [Guid]::NewGuid().ToString("N"))
New-Item -Path $tmp -ItemType Directory -Force | Out-Null

$projPath = Join-Path $tmp "CacheFillRunner.csproj"
$programPath = Join-Path $tmp "Program.cs"

$coreProj = Join-Path $repoRoot "FileExplorer.Core\FileExplorer.Core.csproj"
$infraProj = Join-Path $repoRoot "FileExplorer.Infrastructure\FileExplorer.Infrastructure.csproj"

$projText = @"
<Project Sdk="Microsoft.NET.Sdk">
    <PropertyGroup>
        <OutputType>Exe</OutputType>
        <TargetFramework>net8.0-windows10.0.22621.0</TargetFramework>
    </PropertyGroup>
    <ItemGroup>
        <ProjectReference Include="$coreProj" />
        <ProjectReference Include="$infraProj" />
    </ItemGroup>
</Project>
"@

$programText = @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using FileExplorer.Core.Interfaces;
using FileExplorer.Infrastructure.Services;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging.Abstractions;

if (args.Length < 2)
{
    Console.WriteLine("USAGE_ERROR=Need engine path and at least one root path");
    Environment.Exit(2);
}

var enginePath = args[0];
Environment.SetEnvironmentVariable("FILEEXPLORER_RQ_SEARCH_PATH", enginePath);

var roots = args.Skip(1)
    .Where(path => !string.IsNullOrWhiteSpace(path) && System.IO.Directory.Exists(path))
    .Distinct(StringComparer.OrdinalIgnoreCase)
    .ToList();

if (roots.Count == 0)
{
    Console.WriteLine("ROOTS=0");
    Environment.Exit(3);
}

IIdentityCacheService cache = new SqliteIdentityCacheService(NullLogger<SqliteIdentityCacheService>.Instance);
IIdentityExtractionService extractor = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
IDeepIndexingService deep = new DeepIndexingService(cache, extractor, NullLogger<DeepIndexingService>.Instance);

var sw = Stopwatch.StartNew();
await deep.StartOrRefreshAsync(roots);

var seenRunning = false;
var timeout = TimeSpan.FromSeconds(@TIMEOUT@);
var waitStart = DateTime.UtcNow;

while (DateTime.UtcNow - waitStart < timeout)
{
    if (deep.IsRunning)
    {
        seenRunning = true;
    }

    if (seenRunning && !deep.IsRunning)
    {
        break;
    }

    if (!seenRunning && !string.IsNullOrWhiteSpace(deep.EngineWarning))
    {
        break;
    }

    await Task.Delay(250);
}

sw.Stop();

var rowCount = 0L;
await using (var connection = new SqliteConnection($"Data Source={cache.DatabasePath};Mode=ReadOnly;"))
{
    await connection.OpenAsync();
    await using var command = connection.CreateCommand();
    command.CommandText = "SELECT COUNT(*) FROM FileIdentityCache;";
    var scalar = await command.ExecuteScalarAsync();
    rowCount = Convert.ToInt64(scalar);
}

Console.WriteLine($"ENGINE_WARNING={deep.EngineWarning ?? string.Empty}");
Console.WriteLine($"ROOTS={roots.Count}");
Console.WriteLine($"DB_PATH={cache.DatabasePath}");
Console.WriteLine($"ROW_COUNT={rowCount}");
Console.WriteLine($"ELAPSED_MS={sw.ElapsedMilliseconds}");
"@

$programText = $programText.Replace("@TIMEOUT@", [Math]::Max(30, $TimeoutSeconds).ToString())

Set-Content -Path $projPath -Value $projText -Encoding UTF8
Set-Content -Path $programPath -Value $programText -Encoding UTF8

Push-Location $tmp
try {
    $allArgs = @("run", "--project", $projPath, "--", $resolvedEnginePath) + $favoritesRoots
    dotnet @allArgs
}
finally {
    Pop-Location
    Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Cache fill scan completed."
