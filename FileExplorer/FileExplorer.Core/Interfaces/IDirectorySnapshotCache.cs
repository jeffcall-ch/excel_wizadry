// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// In-memory cache of pre-scanned directory snapshots.
/// Populated by background prefetch at startup so first navigation to a pinned folder is instant.
/// </summary>
public interface IDirectorySnapshotCache
{
    /// <summary>
    /// Stores a snapshot for <paramref name="normalizedPath"/>.
    /// Called by the background prefetcher; overwrites any previous entry.
    /// </summary>
    void Store(string normalizedPath, IReadOnlyList<FileSystemEntry> entries);

    /// <summary>
    /// Attempts to retrieve a pre-scanned snapshot.
    /// Returns <c>true</c> and sets <paramref name="entries"/> when a hit is found;
    /// after the snapshot is consumed it is removed so subsequent navigations re-scan disk.
    /// </summary>
    bool TryConsume(string normalizedPath, out IReadOnlyList<FileSystemEntry> entries);

    /// <summary>Clears the entire cache (e.g. after a FileSystemWatcher change event).</summary>
    void Invalidate(string normalizedPath);
}
