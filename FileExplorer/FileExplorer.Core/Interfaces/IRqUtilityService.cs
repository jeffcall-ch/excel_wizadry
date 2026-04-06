// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Provides rq-backed utility operations beyond search.
/// </summary>
public interface IRqUtilityService
{
    /// <summary>
    /// Computes recursive file count and total size for a folder.
    /// </summary>
    Task<FolderAggregateResult> GetFolderAggregateAsync(string folderPath, CancellationToken cancellationToken = default);
}
