// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Describes a conflict encountered during a file operation.
/// </summary>
public sealed class FileConflict
{
    /// <summary>Source file path causing the conflict.</summary>
    public required string SourcePath { get; init; }

    /// <summary>Destination file path that already exists.</summary>
    public required string DestinationPath { get; init; }

    /// <summary>Size of the source file.</summary>
    public long SourceSize { get; init; }

    /// <summary>Size of the destination file.</summary>
    public long DestinationSize { get; init; }

    /// <summary>Modification date of the source file.</summary>
    public DateTimeOffset SourceDate { get; init; }

    /// <summary>Modification date of the destination file.</summary>
    public DateTimeOffset DestinationDate { get; init; }
}

/// <summary>
/// How to resolve a file conflict.
/// </summary>
public enum ConflictResolution
{
    /// <summary>Skip this file entirely.</summary>
    Skip,

    /// <summary>Replace the existing file.</summary>
    Replace,

    /// <summary>Keep both by renaming the new file with an auto-suffix.</summary>
    RenameWithSuffix,

    /// <summary>Cancel the entire operation.</summary>
    Cancel
}

/// <summary>
/// Result of a conflict resolution including whether to apply to all remaining conflicts.
/// </summary>
public sealed class ConflictResolutionResult
{
    /// <summary>The chosen resolution.</summary>
    public required ConflictResolution Resolution { get; init; }

    /// <summary>Whether to apply this resolution to all remaining conflicts.</summary>
    public bool ApplyToAll { get; init; }
}
