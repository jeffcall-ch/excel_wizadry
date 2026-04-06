// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// rq-backed utility service for recursive folder aggregation.
/// </summary>
public sealed class RqUtilityService : IRqUtilityService
{
    private readonly ILogger<RqUtilityService> _logger;
    private string? _rqPath;

    /// <summary>
    /// Initializes a new instance of the <see cref="RqUtilityService"/> class.
    /// </summary>
    public RqUtilityService(ILogger<RqUtilityService> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public async Task<FolderAggregateResult> GetFolderAggregateAsync(string folderPath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            return new FolderAggregateResult
            {
                IsSuccess = false,
                ErrorMessage = "Folder not found."
            };
        }

        var stopwatch = Stopwatch.StartNew();
        FolderAggregateResult rqResult;

        try
        {
            rqResult = await TryAggregateWithRqAsync(folderPath, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            // rq unavailability or invocation failure should not fail the inspector metadata flow.
            _logger.LogDebug(ex, "rq aggregate unavailable for {Path}. Falling back to directory enumeration.", folderPath);
            rqResult = new FolderAggregateResult
            {
                IsSuccess = false,
                UsedRq = true,
                ErrorMessage = ex.Message
            };
        }

        if (rqResult.IsSuccess)
        {
            stopwatch.Stop();
            return rqResult with { ElapsedMilliseconds = stopwatch.ElapsedMilliseconds };
        }

        try
        {
            var fallbackResult = await Task.Run(() => AggregateWithDirectoryEnumeration(folderPath), cancellationToken);
            stopwatch.Stop();

            if (fallbackResult.IsSuccess)
                return fallbackResult with { ElapsedMilliseconds = stopwatch.ElapsedMilliseconds };

            return fallbackResult with
            {
                ElapsedMilliseconds = stopwatch.ElapsedMilliseconds,
                ErrorMessage = string.IsNullOrWhiteSpace(rqResult.ErrorMessage)
                    ? fallbackResult.ErrorMessage
                    : $"rq: {rqResult.ErrorMessage}; fallback: {fallbackResult.ErrorMessage}"
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogWarning(ex, "Folder aggregate fallback failed for {Path}", folderPath);

            return new FolderAggregateResult
            {
                IsSuccess = false,
                UsedRq = false,
                ElapsedMilliseconds = stopwatch.ElapsedMilliseconds,
                ErrorMessage = string.IsNullOrWhiteSpace(rqResult.ErrorMessage)
                    ? ex.Message
                    : $"rq: {rqResult.ErrorMessage}; fallback exception: {ex.Message}"
            };
        }
    }

    private async Task<FolderAggregateResult> TryAggregateWithRqAsync(string folderPath, CancellationToken cancellationToken)
    {
        var rqPath = GetRqPath();
        var outputLines = new ConcurrentBag<string>();
        var errorOutput = new StringBuilder();

        var processInfo = new ProcessStartInfo
        {
            FileName = rqPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        processInfo.ArgumentList.Add(folderPath);
        processInfo.ArgumentList.Add("*");
        processInfo.ArgumentList.Add("--glob");
        processInfo.ArgumentList.Add("--follow-symlinks");

        using var process = new Process { StartInfo = processInfo };

        process.OutputDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                outputLines.Add(e.Data);
        };

        process.ErrorDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                errorOutput.AppendLine(e.Data);
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw;
        }

        if (process.ExitCode != 0)
        {
            return new FolderAggregateResult
            {
                IsSuccess = false,
                UsedRq = true,
                ErrorMessage = errorOutput.Length > 0
                    ? errorOutput.ToString().Trim()
                    : $"rq exited with code {process.ExitCode}."
            };
        }

        var fileCount = 0;
        long totalBytes = 0;

        foreach (var path in outputLines.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                if (!File.Exists(path))
                    continue;

                var fileInfo = new FileInfo(path);
                fileCount++;
                totalBytes += fileInfo.Length;
            }
            catch
            {
                // Skip inaccessible files.
            }
        }

        return new FolderAggregateResult
        {
            IsSuccess = true,
            UsedRq = true,
            FileCount = fileCount,
            TotalSizeBytes = totalBytes
        };
    }

    private static FolderAggregateResult AggregateWithDirectoryEnumeration(string folderPath)
    {
        var fileCount = 0;
        long totalBytes = 0;

        try
        {
            foreach (var filePath in Directory.EnumerateFiles(folderPath, "*", SearchOption.AllDirectories))
            {
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    fileCount++;
                    totalBytes += fileInfo.Length;
                }
                catch
                {
                    // Skip inaccessible files.
                }
            }

            return new FolderAggregateResult
            {
                IsSuccess = true,
                UsedRq = false,
                FileCount = fileCount,
                TotalSizeBytes = totalBytes
            };
        }
        catch (Exception ex)
        {
            return new FolderAggregateResult
            {
                IsSuccess = false,
                UsedRq = false,
                ErrorMessage = ex.Message
            };
        }
    }

    private string GetRqPath()
    {
        if (_rqPath is not null && File.Exists(_rqPath))
            return _rqPath;

        var localPath = Path.Combine(AppContext.BaseDirectory, "Assets", "rq.exe");
        if (File.Exists(localPath))
        {
            _rqPath = localPath;
            return _rqPath;
        }

        throw new FileNotFoundException("rq.exe not found. Expected at: " + localPath);
    }
}
