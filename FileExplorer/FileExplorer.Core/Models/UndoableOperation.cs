// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents an undoable file operation in the undo stack.
/// </summary>
public sealed class UndoableOperation
{
    /// <summary>Unique identifier for this operation.</summary>
    public Guid Id { get; init; } = Guid.NewGuid();

    /// <summary>Type of the original operation.</summary>
    public required FileOperationType OriginalOperation { get; init; }

    /// <summary>Description of the operation for undo display.</summary>
    public required string Description { get; init; }

    /// <summary>Source paths of the original operation.</summary>
    public required IReadOnlyList<string> OriginalSourcePaths { get; init; }

    /// <summary>Destination paths resulting from the operation.</summary>
    public required IReadOnlyList<string> ResultPaths { get; init; }

    /// <summary>Previous name for rename undo.</summary>
    public string? PreviousName { get; init; }

    /// <summary>Timestamp of the operation.</summary>
    public DateTimeOffset Timestamp { get; init; } = DateTimeOffset.Now;
}
