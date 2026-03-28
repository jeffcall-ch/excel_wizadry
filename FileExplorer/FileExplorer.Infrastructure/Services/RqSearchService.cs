// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using FileExplorer.Core.Interfaces;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Search service that uses the embedded rq.exe fast file search engine.
/// </summary>
public sealed class RqSearchService : ISearchService
{
    private readonly ILogger<RqSearchService> _logger;
    private string? _rqPath;

    public RqSearchService(ILogger<RqSearchService> logger)
    {
        _logger = logger;
    }

    /// <summary>Ensures rq.exe is extracted to a temp location and returns its path.</summary>
    private string GetRqPath()
    {
        if (_rqPath is not null && File.Exists(_rqPath))
            return _rqPath;

        // Look next to the application exe first
        var appDir = AppContext.BaseDirectory;
        var localPath = Path.Combine(appDir, "Assets", "rq.exe");
        if (File.Exists(localPath))
        {
            _rqPath = localPath;
            _logger.LogInformation("Using rq.exe from: {Path}", _rqPath);
            return _rqPath;
        }

        throw new FileNotFoundException("rq.exe not found. Expected at: " + localPath);
    }

    public async Task<IReadOnlyList<string>> SearchAsync(
        string directory, string searchTerm, CancellationToken cancellationToken = default)
    {
        if (!Directory.Exists(directory))
            return Array.Empty<string>();

        var rqPath = GetRqPath();
        var results = new ConcurrentBag<string>();
        var errorOutput = new StringBuilder();

        var psi = new ProcessStartInfo
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

        // Safe argument passing via ArgumentList (no shell injection)
        psi.ArgumentList.Add(directory);
        psi.ArgumentList.Add(searchTerm);
        psi.ArgumentList.Add("--glob");
        psi.ArgumentList.Add("--follow-symlinks");

        _logger.LogInformation("Searching with rq: dir={Dir} term={Term}", directory, searchTerm);

        using var process = new Process { StartInfo = psi };

        process.OutputDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                results.Add(e.Data);
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
            return Array.Empty<string>();
        }

        if (errorOutput.Length > 0)
            _logger.LogWarning("rq stderr: {Error}", errorOutput.ToString());

        _logger.LogInformation("Search completed: {Count} results", results.Count);
        return results.ToList();
    }
}
