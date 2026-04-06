// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// SQLite-backed cache abstraction for extracted document identities.
/// </summary>
public interface IIdentityCacheService
{
    /// <summary>Gets the absolute path to the backing database file.</summary>
    string DatabasePath { get; }

    /// <summary>Normalizes a file or directory path for cache key usage.</summary>
    string NormalizePath(string fullPath);

    /// <summary>Builds the cache key from file path, size, and last-write timestamp.</summary>
    string BuildFileKey(string fullPath, long fileSize, DateTimeOffset lastWriteTime);

    /// <summary>Gets all cached identity rows for a normalized parent folder path.</summary>
    Task<IReadOnlyDictionary<string, FileIdentityCacheRecord>> GetEntriesByParentPathAsync(
        string parentPath,
        CancellationToken cancellationToken = default);

    /// <summary>Gets one cached row by exact file key.</summary>
    Task<FileIdentityCacheRecord?> GetEntryByFileKeyAsync(
        string fileKey,
        CancellationToken cancellationToken = default);

    /// <summary>Upserts cache rows.</summary>
    Task UpsertEntriesAsync(
        IReadOnlyList<FileIdentityCacheRecord> entries,
        CancellationToken cancellationToken = default);

    /// <summary>Deletes rows by exact file keys.</summary>
    Task DeleteByFileKeysAsync(
        IReadOnlyCollection<string> fileKeys,
        CancellationToken cancellationToken = default);

    /// <summary>Deletes all rows for a normalized file path (all key versions).</summary>
    Task DeleteByNormalizedPathAsync(
        string fullPath,
        CancellationToken cancellationToken = default);
}
