// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Service for querying and controlling OneDrive cloud file states via the Cloud Filter API.
/// </summary>
public interface ICloudStatusService
{
    /// <summary>Whether any OneDrive sync root is detected on this machine.</summary>
    bool IsSyncRootDetected { get; }

    /// <summary>Detects installed OneDrive sync roots.</summary>
    Task DetectSyncRootsAsync(CancellationToken cancellationToken = default);

    /// <summary>Queries the cloud status for a single file.</summary>
    Task<CloudFileStatus> GetCloudStatusAsync(string filePath, CancellationToken cancellationToken = default);

    /// <summary>Evaluates cloud status from already-fetched file metadata without additional disk I/O.</summary>
    CloudFileStatus GetCloudStatus(uint fileAttributes, uint reparseTag);

    /// <summary>Queries cloud statuses for a batch of files.</summary>
    Task<IReadOnlyDictionary<string, CloudFileStatus>> GetCloudStatusBatchAsync(IReadOnlyList<string> filePaths, CancellationToken cancellationToken = default);

    /// <summary>Queries cloud statuses for a batch using pre-fetched entry metadata.</summary>
    Task<IReadOnlyDictionary<string, CloudFileStatus>> GetCloudStatusBatchAsync(IReadOnlyList<FileSystemEntry> entries, CancellationToken cancellationToken = default);

    /// <summary>Pins a file for offline availability (downloads it).</summary>
    Task PinFileAsync(string filePath, CancellationToken cancellationToken = default);

    /// <summary>Dehydrates a file to free up local space.</summary>
    Task DehydrateFileAsync(string filePath, CancellationToken cancellationToken = default);

    /// <summary>Returns whether the given path is under an OneDrive sync root.</summary>
    bool IsUnderSyncRoot(string filePath);
}
