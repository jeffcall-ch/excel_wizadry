// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Concurrent;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Thread-safe in-memory store of pre-scanned directory snapshots.
/// Once consumed a snapshot is evicted — live navigation always gets a fresh
/// scan, while the first visit after a prefetch is served from memory.
/// </summary>
public sealed class DirectorySnapshotCache : IDirectorySnapshotCache
{
    // Key: normalized (full) path, OrdinalIgnoreCase
    private readonly ConcurrentDictionary<string, IReadOnlyList<FileSystemEntry>> _store =
        new(StringComparer.OrdinalIgnoreCase);

    /// <inheritdoc/>
    public void Store(string normalizedPath, IReadOnlyList<FileSystemEntry> entries)
        => _store[normalizedPath] = entries;

    /// <inheritdoc/>
    public bool TryConsume(string normalizedPath, out IReadOnlyList<FileSystemEntry> entries)
    {
        if (_store.TryRemove(normalizedPath, out entries!))
            return true;

        entries = [];
        return false;
    }

    /// <inheritdoc/>
    public void Invalidate(string normalizedPath)
        => _store.TryRemove(normalizedPath, out _);
}
