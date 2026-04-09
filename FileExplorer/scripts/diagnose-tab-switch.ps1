<#
.SYNOPSIS
    Measures the cost of a tab switch between two folder paths, mirroring what
    FileExplorer does in the background when ActivateTab() fires.

.DESCRIPTION
    The script measures the following operations in isolation:
      1. Win32 FindFirstFileEx directory enumeration (what FileSystemService does on Task.Run)
      2. UI-thread O(n) post-scan work: dictionary build + HasMaterialChanges check
      3. SQLite cache query for each path (what HydrateIdentityColumnsAsync/RefreshIdentityForSingleFileAsync does)

    Run this script BEFORE deploying a fix to get a baseline, and AFTER to verify.

.EXAMPLE
    pwsh -ExecutionPolicy Bypass -File scripts\diagnose-tab-switch.ps1
#>

param(
    [string]$PathA = "C:\Users\szil\Shortcuts_to_folders",
    [string]$PathB = "C:\Users\szil\OneDrive - Kanadevia Inova\Downloads",
    [string]$AppDir = "$PSScriptRoot\..\artifacts\portable-folder\win-x64\FileExplorer"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Off

# ---------------------------------------------------------------------------
# 0. Helpers
# ---------------------------------------------------------------------------
function Write-Header([string]$text) {
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
}

function Write-Section([string]$text) {
    Write-Host ""
    Write-Host "  -- $text --" -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# 1. Win32 FindFirstFileEx enumeration (identical to FileSystemService.cs)
# ---------------------------------------------------------------------------
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

public static class Win32Enum
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WIN32_FIND_DATA
    {
        public uint   dwFileAttributes;
        public long   ftCreationTime;
        public long   ftLastAccessTime;
        public long   ftLastWriteTime;
        public uint   nFileSizeHigh;
        public uint   nFileSizeLow;
        public uint   dwReserved0;
        public uint   dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]  public string cAlternateFileName;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr FindFirstFileExW(
        string lpFileName, int fInfoLevelId, out WIN32_FIND_DATA lpFindFileData,
        int fSearchOp, IntPtr lpSearchFilter, int dwAdditionalFlags);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool FindNextFileW(IntPtr hFindFile, out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll")]
    private static extern bool FindClose(IntPtr hFindFile);

    private static readonly IntPtr INVALID_HANDLE = new IntPtr(-1);
    private const int FindExInfoBasic   = 1;
    private const int FindExSearchNameMatch = 0;
    private const int FIND_FIRST_EX_LARGE_FETCH = 0x2;

    // Returns: [0]=count  [1]=elapsedMs
    public static long[] EnumerateAndTime(string folder)
    {
        var sw = Stopwatch.StartNew();
        var query = folder.TrimEnd('\\', '/') + @"\*";
        WIN32_FIND_DATA fd;
        var h = FindFirstFileExW(query, FindExInfoBasic, out fd,
                                  FindExSearchNameMatch, IntPtr.Zero,
                                  FIND_FIRST_EX_LARGE_FETCH);
        int count = 0;
        if (h != INVALID_HANDLE)
        {
            try
            {
                do
                {
                    if (fd.cFileName != "." && fd.cFileName != "..")
                        count++;
                }
                while (FindNextFileW(h, out fd));
            }
            finally { FindClose(h); }
        }
        sw.Stop();
        return new long[] { count, sw.ElapsedMilliseconds };
    }
}
'@ -Language CSharp

function Measure-WinEnum {
    param([string]$Path, [int]$Runs = 3)

    Write-Host "    Warming up..." -ForegroundColor DarkGray
    $null = [Win32Enum]::EnumerateAndTime($Path)  # warm-up run (cold read)

    $times = @()
    $count = 0
    for ($i = 1; $i -le $Runs; $i++) {
        $result = [Win32Enum]::EnumerateAndTime($Path)
        $count  = $result[0]
        $times += $result[1]
        Write-Host "      Run $i : $($result[1]) ms  ($($result[0]) items)" -ForegroundColor DarkGray
    }
    $avg = [Math]::Round(($times | Measure-Object -Average).Average, 1)
    return [PSCustomObject]@{ Path = $Path; Count = $count; AvgMs = $avg; Runs = $times }
}

# ---------------------------------------------------------------------------
# 2. SQLite cache stats (identical connection string to SqliteIdentityCacheService)
# ---------------------------------------------------------------------------
function Get-CacheStats {
    param([string]$Path, [string]$DbPath)

    if (-not (Test-Path $DbPath)) {
        return [PSCustomObject]@{ CachedEntries = "DB not found"; DbPath = $DbPath }
    }

    # Try to load Microsoft.Data.Sqlite from the app's published folder
    $sqliteDll = Join-Path $AppDir "Microsoft.Data.Sqlite.dll"
    if (Test-Path $sqliteDll) {
        try {
            Add-Type -Path $sqliteDll -ErrorAction Stop
            Add-Type -Path (Join-Path $AppDir "SQLitePCLRaw.core.dll") -ErrorAction Stop
            Add-Type -Path (Join-Path $AppDir "SQLitePCLRaw.provider.e_sqlite3.dll") -ErrorAction Stop
        } catch {
            # Assemblies already loaded or load not needed
        }

        try {
            $normalized = $Path.ToLowerInvariant().TrimEnd('\', '/')
            $connStr    = "Data Source=$DbPath;Mode=ReadOnly;Cache=Shared"
            $conn       = New-Object Microsoft.Data.Sqlite.SqliteConnection($connStr)
            $conn.Open()

            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT COUNT(*) FROM FileIdentityCache WHERE ParentPath = `$p"
            $null = $cmd.Parameters.AddWithValue('$p', $normalized)
            $cached = $cmd.ExecuteScalar()

            $cmd2 = $conn.CreateCommand()
            $cmd2.CommandText = "SELECT COUNT(*) FROM FileIdentityCache WHERE ParentPath = `$p AND (ProjectTitle IS NOT NULL AND ProjectTitle != '') "
            $null = $cmd2.Parameters.AddWithValue('$p', $normalized)
            $withProject = $cmd2.ExecuteScalar()

            $conn.Close()
            return [PSCustomObject]@{
                CachedEntries    = $cached
                WithProjectTitle = $withProject
                DbPath           = $DbPath
            }
        } catch {
            return [PSCustomObject]@{ CachedEntries = "Query failed: $_"; DbPath = $DbPath }
        }
    } else {
        return [PSCustomObject]@{ CachedEntries = "SQLite DLL not found at $sqliteDll"; DbPath = $DbPath }
    }
}

# Resolve DB path (mirrors SqliteIdentityCacheService.ResolveDatabasePath)
$repoRoot = Split-Path -Parent $PSScriptRoot
$dbPath   = Join-Path $repoRoot "Development\identity_cache.db"
if (-not (Test-Path $dbPath)) {
    $dbPath = Join-Path $env:LOCALAPPDATA "FileExplorer\Development\identity_cache.db"
}

# ---------------------------------------------------------------------------
# 3. Run measurements
# ---------------------------------------------------------------------------
Write-Header "FileExplorer Tab-Switch Diagnostic"
Write-Host "  Path A : $PathA"
Write-Host "  Path B : $PathB"
Write-Host "  DB     : $dbPath"

Write-Section "Step 1 — Win32 directory enumeration (= what Task.Run does)"
Write-Host "  [A] $PathA" -ForegroundColor White
$enumA = Measure-WinEnum -Path $PathA -Runs 3

Write-Host ""
Write-Host "  [B] $PathB" -ForegroundColor White
$enumB = Measure-WinEnum -Path $PathB -Runs 3

Write-Section "Step 2 — UI-thread O(n) post-scan work estimate"
# After Task.Run completes the app builds a Dictionary<string,FSEntry> and runs
# HasMaterialChanges on the UI thread.  Estimate: ~1µs per item.
$uiWorkA = [Math]::Round($enumA.Count * 0.001, 2)
$uiWorkB = [Math]::Round($enumB.Count * 0.001, 2)
Write-Host "  [A] ~$uiWorkA ms   ($($enumA.Count) items × 1 µs)"
Write-Host "  [B] ~$uiWorkB ms   ($($enumB.Count) items × 1 µs)"

Write-Section "Step 3 — SQLite cache status"
$cacheA = Get-CacheStats -Path $PathA -DbPath $dbPath
$cacheB = Get-CacheStats -Path $PathB -DbPath $dbPath
Write-Host "  [A] $PathA"
Write-Host "      Cached entries : $($cacheA.CachedEntries)   (with Project: $($cacheA.WithProjectTitle))"
Write-Host "  [B] $PathB"
Write-Host "      Cached entries : $($cacheB.CachedEntries)   (with Project: $($cacheB.WithProjectTitle))"

Write-Section "Summary"
$fmtA = "{0,4} items  |  avg enum {1,4} ms  |  ui-thread diff ~{2} ms  |  cache {3} entries" `
    -f $enumA.Count, $enumA.AvgMs, $uiWorkA, $cacheA.CachedEntries
$fmtB = "{0,4} items  |  avg enum {1,4} ms  |  ui-thread diff ~{2} ms  |  cache {3} entries" `
    -f $enumB.Count, $enumB.AvgMs, $uiWorkB, $cacheB.CachedEntries

Write-Host "  [A] $fmtA" -ForegroundColor Green
Write-Host "  [B] $fmtB" -ForegroundColor $(if ($enumB.AvgMs -gt 200) { "Red" } elseif ($enumB.AvgMs -gt 50) { "Yellow" } else { "Green" })

Write-Host ""
Write-Host "  Diagnosis:" -ForegroundColor White
if ($enumB.AvgMs -gt 500) {
    Write-Host "  ► Disk enumeration of [B] is SLOW ($($enumB.AvgMs) ms). This is the dominant cost." -ForegroundColor Red
    Write-Host "    This is a OneDrive / network-share latency issue, not a code bug." -ForegroundColor Red
    Write-Host "    Fix: increase the staleness threshold or pre-cache on background without blocking." -ForegroundColor Yellow
} elseif ($enumB.AvgMs -gt 100) {
    Write-Host "  ► Disk enumeration of [B] is MODERATE ($($enumB.AvgMs) ms). Async, should not be felt." -ForegroundColor Yellow
} else {
    Write-Host "  ► Disk enumeration is fast. The bottleneck is likely in the UI thread layout." -ForegroundColor Green
}

if ($enumB.Count -gt 400) {
    Write-Host "  ► High item count ($($enumB.Count)) means the VirtualizingStackPanel layout pass" -ForegroundColor Yellow
    Write-Host "    and ContainerContentChanging callbacks will be proportionally heavier." -ForegroundColor Yellow
}

Write-Host ""
