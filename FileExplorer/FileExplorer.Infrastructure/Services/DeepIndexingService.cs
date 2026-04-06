// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Recursive background indexer for favorites roots using rq_search/rq engine output.
/// </summary>
public sealed class DeepIndexingService : IDeepIndexingService
{
    private readonly IIdentityCacheService _identityCacheService;
    private readonly IIdentityExtractionService _identityExtractionService;
    private readonly ILogger<DeepIndexingService> _logger;
    private readonly SemaphoreSlim _runGate = new(1, 1);

    private CancellationTokenSource? _runCts;

    /// <summary>
    /// Initializes a new instance of the <see cref="DeepIndexingService"/> class.
    /// </summary>
    public DeepIndexingService(
        IIdentityCacheService identityCacheService,
        IIdentityExtractionService identityExtractionService,
        ILogger<DeepIndexingService> logger)
    {
        _identityCacheService = identityCacheService;
        _identityExtractionService = identityExtractionService;
        _logger = logger;
    }

    /// <inheritdoc/>
    public string? EngineWarning { get; private set; }

    /// <inheritdoc/>
    public bool IsRunning { get; private set; }

    /// <inheritdoc/>
    public async Task StartOrRefreshAsync(
        IReadOnlyList<string> favoriteRootPaths,
        CancellationToken cancellationToken = default)
    {
        var validRoots = favoriteRootPaths
            .Where(path => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (validRoots.Count == 0)
            return;

        var enginePath = ResolveRqEnginePath();
        if (string.IsNullOrWhiteSpace(enginePath))
            return;

        await _runGate.WaitAsync(cancellationToken);
        try
        {
            _runCts?.Cancel();
            _runCts?.Dispose();
            _runCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var runToken = _runCts.Token;

            _ = Task.Run(async () =>
            {
                IsRunning = true;
                var originalPriority = Thread.CurrentThread.Priority;
                try
                {
                    Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
                    await RunRecursiveSyncAsync(enginePath, validRoots, runToken).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    // New refresh request superseded this run.
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Deep indexing run failed.");
                }
                finally
                {
                    Thread.CurrentThread.Priority = originalPriority;
                    IsRunning = false;
                }
            }, runToken);
        }
        finally
        {
            _runGate.Release();
        }
    }

    private async Task RunRecursiveSyncAsync(
        string enginePath,
        IReadOnlyList<string> roots,
        CancellationToken cancellationToken)
    {
        var stopwatch = Stopwatch.StartNew();
        var discoveredFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var root in roots)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rootFiles = await DiscoverFilesAsync(enginePath, root, cancellationToken).ConfigureAwait(false);
            foreach (var file in rootFiles)
            {
                discoveredFiles.Add(file);
            }
        }

        var groupedByParent = discoveredFiles
            .Where(path => File.Exists(path))
            .Where(path => !ShouldSkipIndexing(path))
            .GroupBy(path => Path.GetDirectoryName(path) ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .Where(group => !string.IsNullOrWhiteSpace(group.Key))
            .ToList();

        foreach (var group in groupedByParent)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await SyncFolderKeysAsync(group.Key, group.ToList(), cancellationToken).ConfigureAwait(false);
        }

        stopwatch.Stop();
        _logger.LogInformation(
            "Deep indexing finished in {ElapsedMs} ms. Roots={RootCount} Files={FileCount} Folders={FolderCount}",
            stopwatch.ElapsedMilliseconds,
            roots.Count,
            discoveredFiles.Count,
            groupedByParent.Count);
    }

    private async Task<List<string>> DiscoverFilesAsync(string enginePath, string rootPath, CancellationToken cancellationToken)
    {
        var outputLines = new ConcurrentBag<string>();
        var errorOutput = new StringBuilder();

        var processInfo = new ProcessStartInfo
        {
            FileName = enginePath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        processInfo.ArgumentList.Add(rootPath);
        processInfo.ArgumentList.Add("*");
        processInfo.ArgumentList.Add("--glob");
        processInfo.ArgumentList.Add("--follow-symlinks");

        using var process = new Process { StartInfo = processInfo };

        process.OutputDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
                outputLines.Add(args.Data);
        };

        process.ErrorDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
                errorOutput.AppendLine(args.Data);
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw;
        }

        if (process.ExitCode != 0)
        {
            var stderr = errorOutput.ToString().Trim();
            _logger.LogWarning(
                "Deep indexing engine returned non-zero exit code {ExitCode} for root {Root}. stderr={StdErr}",
                process.ExitCode,
                rootPath,
                stderr);
            return [];
        }

        return outputLines
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private async Task SyncFolderKeysAsync(string folderPath, IReadOnlyList<string> filePaths, CancellationToken cancellationToken)
    {
        var normalizedFolder = _identityCacheService.NormalizePath(folderPath);
        if (string.IsNullOrWhiteSpace(normalizedFolder))
            return;

        var liveEntries = new List<(string Path, long Size, DateTimeOffset LastWriteUtc, string Key)>(filePaths.Count);

        foreach (var filePath in filePaths)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var info = new FileInfo(filePath);
                if (!info.Exists)
                    continue;

                var extension = Path.GetExtension(filePath);
                if (!_identityExtractionService.SupportsExtension(extension))
                    continue;

                var key = _identityCacheService.BuildFileKey(filePath, info.Length, info.LastWriteTimeUtc);
                liveEntries.Add((filePath, info.Length, info.LastWriteTimeUtc, key));
            }
            catch
            {
                // Skip inaccessible files; they will be retried in later cycles.
            }
        }

        if (liveEntries.Count == 0)
            return;

        var cacheEntries = await _identityCacheService
            .GetEntriesByParentPathAsync(normalizedFolder, cancellationToken)
            .ConfigureAwait(false);

        var liveKeys = new HashSet<string>(liveEntries.Select(entry => entry.Key), StringComparer.OrdinalIgnoreCase);

        var orphanKeys = cacheEntries.Keys
            .Where(key => !liveKeys.Contains(key))
            .ToList();

        if (orphanKeys.Count > 0)
        {
            await _identityCacheService.DeleteByFileKeysAsync(orphanKeys, cancellationToken).ConfigureAwait(false);
        }

        var missingEntries = liveEntries
            .Where(entry => !cacheEntries.ContainsKey(entry.Key))
            .ToList();

        if (missingEntries.Count == 0)
            return;

        var missingPaths = missingEntries
            .Select(entry => entry.Path)
            .ToList();

        var extracted = await _identityExtractionService
            .ExtractIdentityBatchAsync(missingPaths, cancellationToken)
            .ConfigureAwait(false);

        if (extracted.Count == 0)
            return;

        var upserts = new List<FileIdentityCacheRecord>(missingEntries.Count);
        foreach (var missing in missingEntries)
        {
            if (!extracted.TryGetValue(missing.Path, out var identity) || !identity.HasAnyValue)
                continue;

            upserts.Add(new FileIdentityCacheRecord
            {
                FileKey = missing.Key,
                ParentPath = normalizedFolder,
                ProjectTitle = identity.ProjectTitle,
                Revision = identity.Revision,
                DocumentName = identity.DocumentName,
                LastWriteTime = missing.LastWriteUtc.UtcTicks,
                FileSize = missing.Size
            });
        }

        if (upserts.Count > 0)
        {
            await _identityCacheService.UpsertEntriesAsync(upserts, cancellationToken).ConfigureAwait(false);
        }
    }

    private string? ResolveRqEnginePath()
    {
        foreach (var candidate in GetEnginePathCandidates())
        {
            if (string.IsNullOrWhiteSpace(candidate))
                continue;

            var expanded = Environment.ExpandEnvironmentVariables(candidate.Trim());
            if (File.Exists(expanded))
            {
                EngineWarning = null;
                return expanded;
            }
        }

        EngineWarning = "rq_search.exe not found. Set Config.json RqSearchPath or FILEEXPLORER_RQ_SEARCH_PATH.";
        return null;
    }

    private IEnumerable<string> GetEnginePathCandidates()
    {
        var fromConfig = ReadEnginePathFromConfig();
        if (!string.IsNullOrWhiteSpace(fromConfig))
            yield return fromConfig;

        var fromEnvironment = Environment.GetEnvironmentVariable("FILEEXPLORER_RQ_SEARCH_PATH");
        if (!string.IsNullOrWhiteSpace(fromEnvironment))
            yield return fromEnvironment;

        yield return Path.Combine(AppContext.BaseDirectory, "Assets", "rq_search.exe");
        yield return Path.Combine(AppContext.BaseDirectory, "Assets", "rq.exe");
    }

    private static string? ReadEnginePathFromConfig()
    {
        foreach (var configPath in GetConfigPathCandidates())
        {
            if (!File.Exists(configPath))
                continue;

            try
            {
                using var stream = File.OpenRead(configPath);
                using var document = JsonDocument.Parse(stream);

                if (document.RootElement.ValueKind != JsonValueKind.Object)
                    continue;

                if (document.RootElement.TryGetProperty("RqSearchPath", out var pathNode) &&
                    pathNode.ValueKind == JsonValueKind.String)
                {
                    return pathNode.GetString();
                }

                if (document.RootElement.TryGetProperty("RqSearchExePath", out var exePathNode) &&
                    exePathNode.ValueKind == JsonValueKind.String)
                {
                    return exePathNode.GetString();
                }
            }
            catch
            {
                // Ignore malformed config and continue fallback chain.
            }
        }

        return null;
    }

    private static IEnumerable<string> GetConfigPathCandidates()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "Config.json");

        var current = new DirectoryInfo(AppContext.BaseDirectory);
        for (var i = 0; i < 12 && current is not null; i++)
        {
            yield return Path.Combine(current.FullName, "Config.json");
            current = current.Parent;
        }
    }

    private static bool ShouldSkipIndexing(string filePath)
    {
        try
        {
            var fileName = Path.GetFileName(filePath);
            if (fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                return true;

            var attributes = File.GetAttributes(filePath);
            if ((attributes & FileAttributes.Hidden) != 0)
                return true;

            if ((attributes & FileAttributes.System) != 0)
                return true;

            return false;
        }
        catch
        {
            return true;
        }
    }
}
