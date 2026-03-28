// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Service for file system operations including async enumeration, the operation queue,
/// file system watching, and undo support.
/// </summary>
public interface IFileSystemService : IDisposable
{
    /// <summary>Raised when a conflict occurs during an operation.</summary>
    event Func<FileConflict, Task<ConflictResolutionResult>>? ConflictResolutionRequested;

    /// <summary>Raised when an operation encounters an error.</summary>
    event EventHandler<FileOperationErrorEventArgs>? OperationError;

    /// <summary>Enumerates directory contents asynchronously.</summary>
    IAsyncEnumerable<FileSystemEntry> EnumerateDirectoryAsync(string path, CancellationToken cancellationToken = default);

    /// <summary>Enqueues a file operation onto the background processor.</summary>
    Task EnqueueOperationAsync(FileOperation operation, IProgress<FileOperationProgress>? progress = null, CancellationToken cancellationToken = default);

    /// <summary>Creates a new folder at the specified path, handling name collisions.</summary>
    Task<string> CreateNewFolderAsync(string parentPath, CancellationToken cancellationToken = default);

    /// <summary>Renames a file or directory.</summary>
    Task<bool> RenameAsync(string fullPath, string newName, CancellationToken cancellationToken = default);

    /// <summary>Deletes items to the Recycle Bin.</summary>
    Task<bool> DeleteToRecycleBinAsync(IReadOnlyList<string> paths, CancellationToken cancellationToken = default);

    /// <summary>Permanently deletes items.</summary>
    Task<bool> PermanentDeleteAsync(IReadOnlyList<string> paths, CancellationToken cancellationToken = default);

    /// <summary>Undoes the last operation.</summary>
    Task<bool> UndoLastOperationAsync(CancellationToken cancellationToken = default);

    /// <summary>Whether there are operations to undo.</summary>
    bool CanUndo { get; }

    /// <summary>Starts watching a directory for changes.</summary>
    IDisposable WatchDirectory(string path, Action<FileSystemChangeEvent> onChange);

    /// <summary>Gets the current operation queue depth.</summary>
    int QueueDepth { get; }
}

/// <summary>
/// Event args for file operation errors.
/// </summary>
public sealed class FileOperationErrorEventArgs : EventArgs
{
    /// <summary>The full path of the file that caused the error.</summary>
    public required string FilePath { get; init; }

    /// <summary>The error message.</summary>
    public required string ErrorMessage { get; init; }

    /// <summary>The Win32 error code.</summary>
    public int Win32ErrorCode { get; init; }

    /// <summary>Whether to offer "Retry as Administrator".</summary>
    public bool IsAccessDenied { get; init; }
}

/// <summary>
/// Describes a file system change event (debounced).
/// </summary>
public sealed class FileSystemChangeEvent
{
    /// <summary>Type of change.</summary>
    public required WatcherChangeTypes ChangeType { get; init; }

    /// <summary>Full path of the changed item.</summary>
    public required string FullPath { get; init; }

    /// <summary>Previous full path (for renames).</summary>
    public string? OldFullPath { get; init; }

    /// <summary>Display name of the changed item.</summary>
    public required string Name { get; init; }
}
