// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Extracts identity fields from supported document formats.
/// </summary>
public interface IIdentityExtractionService
{
    /// <summary>
    /// Returns true when the provided extension is supported by the extraction pipeline.
    /// </summary>
    bool SupportsExtension(string extension);

    /// <summary>
    /// Extracts identity fields for a single file.
    /// </summary>
    Task<DocumentIdentity> ExtractIdentityAsync(string filePath, CancellationToken cancellationToken = default);

    /// <summary>
    /// Extracts identity fields for multiple files in parallel.
    /// </summary>
    Task<IReadOnlyDictionary<string, DocumentIdentity>> ExtractIdentityBatchAsync(
        IReadOnlyList<string> filePaths,
        CancellationToken cancellationToken = default);
}
