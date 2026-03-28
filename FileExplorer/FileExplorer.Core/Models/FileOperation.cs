// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Describes a queued file operation (copy, move, delete, rename).
/// </summary>
public sealed class FileOperation
{
    /// <summary>Unique identifier for this operation.</summary>
    public Guid Id { get; init; } = Guid.NewGuid();

    /// <summary>Type of operation to perform.</summary>
    public required FileOperationType OperationType { get; init; }

    /// <summary>Source paths for the operation.</summary>
    public required IReadOnlyList<string> SourcePaths { get; init; }

    /// <summary>Destination path (for Copy/Move). Null for Delete.</summary>
    public string? DestinationPath { get; init; }

    /// <summary>New name for rename operations.</summary>
    public string? NewName { get; init; }

    /// <summary>Whether to permanently delete (skip Recycle Bin).</summary>
    public bool PermanentDelete { get; init; }

    /// <summary>Cancellation token source for this operation batch.</summary>
    public CancellationTokenSource CancellationSource { get; } = new();
}

/// <summary>
/// Types of file system operations.
/// </summary>
public enum FileOperationType
{
    Copy,
    Move,
    Delete,
    Rename,
    CreateFolder
}
