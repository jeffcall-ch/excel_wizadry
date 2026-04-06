// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Recursive aggregation result for a folder.
/// </summary>
public sealed record FolderAggregateResult
{
    /// <summary>Total files found recursively.</summary>
    public int FileCount { get; init; }

    /// <summary>Total bytes across all files found recursively.</summary>
    public long TotalSizeBytes { get; init; }

    /// <summary>Operation elapsed time in milliseconds.</summary>
    public long ElapsedMilliseconds { get; init; }

    /// <summary>True when rq utility path was used.</summary>
    public bool UsedRq { get; init; }

    /// <summary>True when the operation completed successfully.</summary>
    public bool IsSuccess { get; init; }

    /// <summary>Error message when operation fails.</summary>
    public string ErrorMessage { get; init; } = string.Empty;
}
