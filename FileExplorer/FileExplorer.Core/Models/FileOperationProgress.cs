// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Progress report for an ongoing file operation.
/// </summary>
public sealed class FileOperationProgress
{
    /// <summary>Operation ID this progress belongs to.</summary>
    public Guid OperationId { get; init; }

    /// <summary>Current file being processed.</summary>
    public string CurrentFile { get; init; } = string.Empty;

    /// <summary>Total number of items in the operation.</summary>
    public int TotalItems { get; init; }

    /// <summary>Number of items completed so far.</summary>
    public int CompletedItems { get; init; }

    /// <summary>Total bytes to process.</summary>
    public long TotalBytes { get; init; }

    /// <summary>Bytes processed so far.</summary>
    public long BytesCompleted { get; init; }

    /// <summary>Current transfer speed in bytes per second.</summary>
    public double BytesPerSecond { get; init; }

    /// <summary>Estimated time remaining.</summary>
    public TimeSpan? EstimatedTimeRemaining { get; init; }

    /// <summary>Overall progress percentage (0-100).</summary>
    public double ProgressPercent => TotalBytes > 0 ? (double)BytesCompleted / TotalBytes * 100.0 : 0;

    /// <summary>Whether the operation completed.</summary>
    public bool IsComplete { get; init; }

    /// <summary>Whether the operation failed.</summary>
    public bool IsFailed { get; init; }

    /// <summary>Error message if operation failed.</summary>
    public string? ErrorMessage { get; init; }
}
