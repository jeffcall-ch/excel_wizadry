// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Channels;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Infrastructure.Native;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Production implementation of <see cref="IFileSystemService"/> backed by Win32 APIs,
/// a Channel-based operation queue, and FileSystemWatcher with debounced change events.
/// </summary>
public sealed class FileSystemService : IFileSystemService
{
    private const int ErrorAccessDenied = 5;
    private const int ErrorSharingViolation = 32;
    private const int ErrorLockViolation = 33;

    private readonly ILogger<FileSystemService> _logger;
    private readonly Channel<QueuedOperation> _operationChannel;
    private readonly ConcurrentStack<UndoableOperation> _undoStack = new();
    private readonly CancellationTokenSource _serviceCts = new();
    private readonly Task _processorTask;
    private int _queueDepth;

    /// <inheritdoc/>
    public event Func<FileConflict, Task<ConflictResolutionResult>>? ConflictResolutionRequested;

    /// <inheritdoc/>
    public event EventHandler<FileOperationErrorEventArgs>? OperationError;

    /// <inheritdoc/>
    public bool CanUndo => !_undoStack.IsEmpty;

    /// <inheritdoc/>
    public int QueueDepth => _queueDepth;

    /// <summary>Initializes a new instance of the <see cref="FileSystemService"/> class.</summary>
    public FileSystemService(ILogger<FileSystemService> logger)
    {
        _logger = logger;
        _operationChannel = Channel.CreateBounded<QueuedOperation>(new BoundedChannelOptions(256)
        {
            FullMode = BoundedChannelFullMode.Wait,
            SingleReader = true,
            SingleWriter = false
        });
        _processorTask = Task.Run(() => ProcessOperationQueueAsync(_serviceCts.Token));
    }

    #region Directory Enumeration

    /// <inheritdoc/>
    public async IAsyncEnumerable<FileSystemEntry> EnumerateDirectoryAsync(
        string path, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await Task.CompletedTask; // Preserve async iterator semantics.

        var normalizedPath = Path.GetFullPath(path);
        var queryPath = Path.Combine(NativeMethods.EnsureLongPath(normalizedPath), "*");

        nint findHandle = NativeMethods.FindFirstFileEx(
            queryPath,
            NativeMethods.FindExInfoBasic,
            out var findData,
            NativeMethods.FindExSearchNameMatch,
            nint.Zero,
            NativeMethods.FIND_FIRST_EX_LARGE_FETCH);

        if (findHandle == NativeMethods.INVALID_HANDLE_VALUE)
        {
            var error = Marshal.GetLastPInvokeError();
            if (error == NativeMethods.ERROR_NO_MORE_FILES)
                yield break;

            var message = NativeMethods.FormatWin32Error(error);
            _logger.LogWarning("Failed to enumerate {Path}: {Error}", path, message);
            RaiseOperationError(path, message, error, isAccessDenied: error == 5);
            yield break;
        }

        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var name = findData.cFileName;
                if (!string.IsNullOrEmpty(name) &&
                    !string.Equals(name, ".", StringComparison.Ordinal) &&
                    !string.Equals(name, "..", StringComparison.Ordinal))
                {
                    var attributes = (FileAttributes)findData.dwFileAttributes;
                    var isDir = (attributes & FileAttributes.Directory) != 0;
                    var fullPath = Path.Combine(normalizedPath, name);
                    var size = isDir ? 0L : ((long)findData.nFileSizeHigh << 32) | findData.nFileSizeLow;
                    var reparseTag = (findData.dwFileAttributes & NativeMethods.FILE_ATTRIBUTE_REPARSE_POINT) != 0
                        ? findData.dwReserved0
                        : 0u;

                    yield return new FileSystemEntry
                    {
                        FullPath = fullPath,
                        Name = name,
                        IsDirectory = isDir,
                        Size = size,
                        DateModified = FileTimeToDateTimeOffset(findData.ftLastWriteTime),
                        Extension = isDir ? string.Empty : Path.GetExtension(name),
                        Attributes = attributes,
                        ReparseTag = reparseTag
                    };
                }

                if (!NativeMethods.FindNextFile(findHandle, out findData))
                {
                    var error = Marshal.GetLastPInvokeError();
                    if (error == NativeMethods.ERROR_NO_MORE_FILES)
                        break;

                    _logger.LogDebug("FindNextFile failed for {Path}: {Error}", path, NativeMethods.FormatWin32Error(error));
                    break;
                }
            }
        }
        finally
        {
            NativeMethods.FindClose(findHandle);
        }
    }

    #endregion

    #region Operation Queue

    /// <inheritdoc/>
    public async Task EnqueueOperationAsync(
        FileOperation operation,
        IProgress<FileOperationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var queued = new QueuedOperation(operation, progress);
        Interlocked.Increment(ref _queueDepth);
        await _operationChannel.Writer.WriteAsync(queued, cancellationToken);
    }

    private async Task ProcessOperationQueueAsync(CancellationToken cancellationToken)
    {
        await foreach (var queued in _operationChannel.Reader.ReadAllAsync(cancellationToken))
        {
            try
            {
                var op = queued.Operation;
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, op.CancellationSource.Token);

                switch (op.OperationType)
                {
                    case FileOperationType.Copy:
                        await ExecuteCopyAsync(op, queued.Progress, linkedCts.Token);
                        break;
                    case FileOperationType.Move:
                        await ExecuteMoveAsync(op, queued.Progress, linkedCts.Token);
                        break;
                    case FileOperationType.Delete:
                        await ExecuteDeleteAsync(op, queued.Progress, linkedCts.Token);
                        break;
                    case FileOperationType.Rename:
                        await ExecuteRenameOperationAsync(op, queued.Progress, linkedCts.Token);
                        break;
                    case FileOperationType.CreateFolder:
                        await ExecuteCreateFolderAsync(op, queued.Progress, linkedCts.Token);
                        break;
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogInformation("Operation {Id} was cancelled", queued.Operation.Id);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Operation {Id} failed", queued.Operation.Id);
                queued.Progress?.Report(new FileOperationProgress
                {
                    OperationId = queued.Operation.Id,
                    IsFailed = true,
                    ErrorMessage = ex.Message
                });
            }
            finally
            {
                Interlocked.Decrement(ref _queueDepth);
            }
        }
    }

    private async Task ExecuteCopyAsync(FileOperation op, IProgress<FileOperationProgress>? progress, CancellationToken ct)
    {
        if (op.DestinationPath is null) return;
        var sw = Stopwatch.StartNew();
        var resultPaths = new List<string>();
        long totalBytes = 0;
        long copiedBytes = 0;

        foreach (var src in op.SourcePaths)
        {
            if (File.Exists(src))
                totalBytes += new FileInfo(src).Length;
        }

        for (int i = 0; i < op.SourcePaths.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var src = op.SourcePaths[i];
            var destName = Path.GetFileName(src);
            var destPath = Path.Combine(op.DestinationPath, destName);

            var isSameFolderCopy = string.Equals(
                NormalizeDirectoryPath(Path.GetDirectoryName(src)),
                NormalizeDirectoryPath(op.DestinationPath),
                StringComparison.OrdinalIgnoreCase);

            if (isSameFolderCopy && (File.Exists(destPath) || Directory.Exists(destPath)))
            {
                destPath = GenerateExplorerCopyPath(destPath);
            }
            else
            {
                destPath = await ResolveConflictIfNeededAsync(src, destPath, ct);
            }

            if (destPath is null) continue; // Skipped

            await Task.Run(() =>
            {
                if (Directory.Exists(src))
                    CopyDirectoryRecursive(src, destPath, ct);
                else
                    File.Copy(NativeMethods.EnsureLongPath(src), NativeMethods.EnsureLongPath(destPath), overwrite: true);
            }, ct);

            resultPaths.Add(destPath);
            if (File.Exists(src)) copiedBytes += new FileInfo(src).Length;

            progress?.Report(new FileOperationProgress
            {
                OperationId = op.Id,
                CurrentFile = destName,
                TotalItems = op.SourcePaths.Count,
                CompletedItems = i + 1,
                TotalBytes = totalBytes,
                BytesCompleted = copiedBytes,
                BytesPerSecond = sw.Elapsed.TotalSeconds > 0 ? copiedBytes / sw.Elapsed.TotalSeconds : 0,
                EstimatedTimeRemaining = copiedBytes > 0 ? TimeSpan.FromSeconds((totalBytes - copiedBytes) / (copiedBytes / sw.Elapsed.TotalSeconds)) : null
            });
        }

        _undoStack.Push(new UndoableOperation
        {
            OriginalOperation = FileOperationType.Copy,
            Description = $"Copy {op.SourcePaths.Count} item(s)",
            OriginalSourcePaths = op.SourcePaths,
            ResultPaths = resultPaths
        });

        progress?.Report(new FileOperationProgress { OperationId = op.Id, IsComplete = true, TotalItems = op.SourcePaths.Count, CompletedItems = op.SourcePaths.Count });
    }

    private async Task ExecuteMoveAsync(FileOperation op, IProgress<FileOperationProgress>? progress, CancellationToken ct)
    {
        if (op.DestinationPath is null) return;
        var resultPaths = new List<string>();

        for (int i = 0; i < op.SourcePaths.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var src = op.SourcePaths[i];
            var destName = Path.GetFileName(src);
            var destPath = Path.Combine(op.DestinationPath, destName);

            destPath = await ResolveConflictIfNeededAsync(src, destPath, ct);
            if (destPath is null) continue;

            await Task.Run(() =>
            {
                var longSrc = NativeMethods.EnsureLongPath(src);
                var longDest = NativeMethods.EnsureLongPath(destPath);
                if (!NativeMethods.MoveFileEx(longSrc, longDest, NativeMethods.MOVEFILE_COPY_ALLOWED))
                {
                    int err = Marshal.GetLastPInvokeError();
                    throw new IOException($"MoveFileEx failed: {NativeMethods.FormatWin32Error(err)}");
                }
            }, ct);

            resultPaths.Add(destPath);
            progress?.Report(new FileOperationProgress
            {
                OperationId = op.Id,
                CurrentFile = destName,
                TotalItems = op.SourcePaths.Count,
                CompletedItems = i + 1,
            });
        }

        _undoStack.Push(new UndoableOperation
        {
            OriginalOperation = FileOperationType.Move,
            Description = $"Move {op.SourcePaths.Count} item(s)",
            OriginalSourcePaths = op.SourcePaths,
            ResultPaths = resultPaths
        });

        progress?.Report(new FileOperationProgress { OperationId = op.Id, IsComplete = true });
    }

    private async Task ExecuteDeleteAsync(FileOperation op, IProgress<FileOperationProgress>? progress, CancellationToken ct)
    {
        if (op.PermanentDelete)
        {
            for (int i = 0; i < op.SourcePaths.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                await Task.Run(() =>
                {
                    var longPath = NativeMethods.EnsureLongPath(op.SourcePaths[i]);
                    if (Directory.Exists(longPath))
                        Directory.Delete(longPath, recursive: true);
                    else if (File.Exists(longPath))
                        File.Delete(longPath);
                }, ct);

                progress?.Report(new FileOperationProgress
                {
                    OperationId = op.Id,
                    CurrentFile = Path.GetFileName(op.SourcePaths[i]),
                    TotalItems = op.SourcePaths.Count,
                    CompletedItems = i + 1
                });
            }
        }
        else
        {
            await Task.Run(() => SendToRecycleBin(op.SourcePaths), ct);
        }

        progress?.Report(new FileOperationProgress { OperationId = op.Id, IsComplete = true });
    }

    private async Task ExecuteRenameOperationAsync(FileOperation op, IProgress<FileOperationProgress>? progress, CancellationToken ct)
    {
        if (op.SourcePaths.Count == 0 || op.NewName is null) return;
        var src = op.SourcePaths[0];
        var dir = Path.GetDirectoryName(src)!;
        var dest = Path.Combine(dir, op.NewName);

        await Task.Run(() =>
        {
            if (!NativeMethods.MoveFileEx(
                NativeMethods.EnsureLongPath(src),
                NativeMethods.EnsureLongPath(dest),
                NativeMethods.MOVEFILE_WRITE_THROUGH))
            {
                int err = Marshal.GetLastPInvokeError();
                throw new IOException($"Rename failed: {NativeMethods.FormatWin32Error(err)}");
            }
        }, ct);

        _undoStack.Push(new UndoableOperation
        {
            OriginalOperation = FileOperationType.Rename,
            Description = $"Rename {Path.GetFileName(src)} to {op.NewName}",
            OriginalSourcePaths = [src],
            ResultPaths = [dest],
            PreviousName = Path.GetFileName(src)
        });

        progress?.Report(new FileOperationProgress { OperationId = op.Id, IsComplete = true });
    }

    private async Task ExecuteCreateFolderAsync(FileOperation op, IProgress<FileOperationProgress>? progress, CancellationToken ct)
    {
        if (op.DestinationPath is null) return;
        await Task.Run(() => Directory.CreateDirectory(NativeMethods.EnsureLongPath(op.DestinationPath)), ct);

        _undoStack.Push(new UndoableOperation
        {
            OriginalOperation = FileOperationType.CreateFolder,
            Description = $"Create folder {Path.GetFileName(op.DestinationPath)}",
            OriginalSourcePaths = [],
            ResultPaths = [op.DestinationPath]
        });

        progress?.Report(new FileOperationProgress { OperationId = op.Id, IsComplete = true });
    }

    #endregion

    #region Public API Methods

    /// <inheritdoc/>
    public async Task<string> CreateNewFolderAsync(string parentPath, CancellationToken cancellationToken = default)
    {
        string newFolderName = "New folder";
        string newFolderPath = Path.Combine(parentPath, newFolderName);
        int counter = 2;

        while (Directory.Exists(newFolderPath) || File.Exists(newFolderPath))
        {
            newFolderName = $"New folder ({counter++})";
            newFolderPath = Path.Combine(parentPath, newFolderName);
        }

        await Task.Run(() => Directory.CreateDirectory(NativeMethods.EnsureLongPath(newFolderPath)), cancellationToken);

        _undoStack.Push(new UndoableOperation
        {
            OriginalOperation = FileOperationType.CreateFolder,
            Description = $"Create folder {newFolderName}",
            OriginalSourcePaths = [],
            ResultPaths = [newFolderPath]
        });

        _logger.LogInformation("Created new folder: {Path}", newFolderPath);
        return newFolderPath;
    }

    /// <inheritdoc/>
    public async Task<bool> RenameAsync(string fullPath, string newName, CancellationToken cancellationToken = default)
    {
        var dir = Path.GetDirectoryName(fullPath)!;
        var newPath = Path.Combine(dir, newName);
        var oldName = Path.GetFileName(fullPath);
        int errorCode = 0;

        try
        {
            await Task.Run(() =>
            {
                if (!NativeMethods.MoveFileEx(
                    NativeMethods.EnsureLongPath(fullPath),
                    NativeMethods.EnsureLongPath(newPath),
                    NativeMethods.MOVEFILE_WRITE_THROUGH))
                {
                    errorCode = Marshal.GetLastPInvokeError();
                    throw new IOException(
                        NativeMethods.FormatWin32Error(errorCode),
                        new Win32Exception(errorCode));
                }
            }, cancellationToken);

            _undoStack.Push(new UndoableOperation
            {
                OriginalOperation = FileOperationType.Rename,
                Description = $"Rename {oldName} to {newName}",
                OriginalSourcePaths = [fullPath],
                ResultPaths = [newPath],
                PreviousName = oldName
            });

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to rename {Path} to {NewName}", fullPath, newName);
            if (errorCode == 0)
                errorCode = GetWin32ErrorCode(ex);

            RaiseOperationError(
                fullPath,
                BuildOperationErrorMessage(ex, errorCode),
                errorCode,
                errorCode == ErrorAccessDenied);
            return false;
        }
    }

    /// <inheritdoc/>
    public Task<bool> DeleteToRecycleBinAsync(IReadOnlyList<string> paths, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            var firstPath = paths.FirstOrDefault() ?? string.Empty;

            try
            {
                foreach (var path in paths)
                {
                    if (TryDetectInUseFile(path, out var inUseError))
                    {
                        RaiseOperationError(path, NativeMethods.FormatWin32Error(inUseError), inUseError, isAccessDenied: false);
                        return false;
                    }
                }

                SendToRecycleBin(paths);
                return true;
            }
            catch (Exception ex)
            {
                var errorCode = GetWin32ErrorCode(ex);
                _logger.LogError(ex, "Failed to send items to Recycle Bin");
                RaiseOperationError(
                    firstPath,
                    BuildOperationErrorMessage(ex, errorCode),
                    errorCode,
                    errorCode == ErrorAccessDenied);
                return false;
            }
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public Task<bool> PermanentDeleteAsync(IReadOnlyList<string> paths, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            var currentPath = paths.FirstOrDefault() ?? string.Empty;

            try
            {
                foreach (var path in paths)
                {
                    currentPath = path;
                    cancellationToken.ThrowIfCancellationRequested();
                    var longPath = NativeMethods.EnsureLongPath(path);
                    if (Directory.Exists(longPath))
                        Directory.Delete(longPath, recursive: true);
                    else if (File.Exists(longPath))
                        File.Delete(longPath);
                }
                return true;
            }
            catch (Exception ex)
            {
                var errorCode = GetWin32ErrorCode(ex);
                _logger.LogError(ex, "Failed to permanently delete items");
                RaiseOperationError(
                    currentPath,
                    BuildOperationErrorMessage(ex, errorCode),
                    errorCode,
                    errorCode == ErrorAccessDenied);
                return false;
            }
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public async Task<bool> UndoLastOperationAsync(CancellationToken cancellationToken = default)
    {
        if (!_undoStack.TryPop(out var undoOp)) return false;

        try
        {
            switch (undoOp.OriginalOperation)
            {
                case FileOperationType.Rename:
                    if (undoOp.ResultPaths.Count > 0 && undoOp.PreviousName is not null)
                    {
                        var dir = Path.GetDirectoryName(undoOp.ResultPaths[0])!;
                        var oldPath = Path.Combine(dir, undoOp.PreviousName);
                        await Task.Run(() => NativeMethods.MoveFileEx(
                            NativeMethods.EnsureLongPath(undoOp.ResultPaths[0]),
                            NativeMethods.EnsureLongPath(oldPath),
                            NativeMethods.MOVEFILE_WRITE_THROUGH), cancellationToken);
                    }
                    break;

                case FileOperationType.Copy:
                    foreach (var path in undoOp.ResultPaths)
                    {
                        await Task.Run(() =>
                        {
                            if (Directory.Exists(path)) Directory.Delete(path, true);
                            else if (File.Exists(path)) File.Delete(path);
                        }, cancellationToken);
                    }
                    break;

                case FileOperationType.Move:
                    for (int i = 0; i < undoOp.ResultPaths.Count && i < undoOp.OriginalSourcePaths.Count; i++)
                    {
                        var result = undoOp.ResultPaths[i];
                        var original = undoOp.OriginalSourcePaths[i];
                        await Task.Run(() => NativeMethods.MoveFileEx(
                            NativeMethods.EnsureLongPath(result),
                            NativeMethods.EnsureLongPath(original),
                            NativeMethods.MOVEFILE_COPY_ALLOWED | NativeMethods.MOVEFILE_WRITE_THROUGH), cancellationToken);
                    }
                    break;

                case FileOperationType.CreateFolder:
                    foreach (var path in undoOp.ResultPaths)
                    {
                        await Task.Run(() =>
                        {
                            if (Directory.Exists(path)) Directory.Delete(path, false);
                        }, cancellationToken);
                    }
                    break;
            }

            _logger.LogInformation("Undid operation: {Description}", undoOp.Description);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to undo operation: {Description}", undoOp.Description);
            _undoStack.Push(undoOp); // Re-push on failure
            return false;
        }
    }

    #endregion

    #region FileSystemWatcher

    /// <inheritdoc/>
    public IDisposable WatchDirectory(string path, Action<FileSystemChangeEvent> onChange)
    {
        var watcher = new FileSystemWatcher(path)
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName,
            IncludeSubdirectories = false,
            EnableRaisingEvents = true
        };

        var debouncer = new ChangeDebouncer(onChange, TimeSpan.FromMilliseconds(250));

        watcher.Created += (_, e) => debouncer.Enqueue(new FileSystemChangeEvent { ChangeType = WatcherChangeTypes.Created, FullPath = e.FullPath, Name = e.Name ?? Path.GetFileName(e.FullPath) });
        watcher.Deleted += (_, e) => debouncer.Enqueue(new FileSystemChangeEvent { ChangeType = WatcherChangeTypes.Deleted, FullPath = e.FullPath, Name = e.Name ?? Path.GetFileName(e.FullPath) });
        watcher.Renamed += (_, e) => debouncer.Enqueue(new FileSystemChangeEvent { ChangeType = WatcherChangeTypes.Renamed, FullPath = e.FullPath, OldFullPath = e.OldFullPath, Name = e.Name ?? Path.GetFileName(e.FullPath) });
        // Changed notifications are intentionally ignored to avoid bursty UI churn from metadata-only updates.

        watcher.Error += (_, e) =>
        {
            _logger.LogWarning(e.GetException(), "FileSystemWatcher error for {Path}. Buffer overflow — falling back to scheduled refresh.", path);
        };

        return new WatcherHandle(watcher, debouncer);
    }

    #endregion

    #region Helpers

    private void SendToRecycleBin(IReadOnlyList<string> paths)
    {
        // SHFileOperation requires double null-terminated string
        var from = string.Join('\0', paths) + "\0\0";
        var fileOp = new NativeMethods.SHFILEOPSTRUCT
        {
            wFunc = NativeMethods.FO_DELETE,
            pFrom = from,
            fFlags = NativeMethods.FOF_ALLOWUNDO | NativeMethods.FOF_NOCONFIRMATION | NativeMethods.FOF_SILENT | NativeMethods.FOF_NOERRORUI
        };

        int result = NativeMethods.SHFileOperation(ref fileOp);
        if (result != 0)
        {
            throw new Win32Exception(result);
        }
    }

    private static bool TryDetectInUseFile(string path, out int errorCode)
    {
        errorCode = 0;

        if (!File.Exists(path))
            return false;

        try
        {
            using var _ = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException ex)
        {
            errorCode = GetWin32ErrorCode(ex);
            return errorCode is ErrorSharingViolation or ErrorLockViolation;
        }
        catch
        {
            return false;
        }
    }

    private static int GetWin32ErrorCode(Exception ex)
    {
        if (ex is Win32Exception win32 && win32.NativeErrorCode > 0)
            return win32.NativeErrorCode;

        if (ex.InnerException is Win32Exception innerWin32 && innerWin32.NativeErrorCode > 0)
            return innerWin32.NativeErrorCode;

        return ex.HResult is 0 ? 0 : ex.HResult & 0xFFFF;
    }

    private static string BuildOperationErrorMessage(Exception ex, int errorCode)
    {
        if (errorCode > 0)
            return NativeMethods.FormatWin32Error(errorCode);

        return ex.Message;
    }

    private async Task<string?> ResolveConflictIfNeededAsync(string sourcePath, string destPath, CancellationToken ct)
    {
        if (!File.Exists(destPath) && !Directory.Exists(destPath))
            return destPath;

        if (ConflictResolutionRequested is null)
            return GenerateAutoSuffixPath(destPath);

        var conflict = new FileConflict
        {
            SourcePath = sourcePath,
            DestinationPath = destPath,
            SourceSize = File.Exists(sourcePath) ? new FileInfo(sourcePath).Length : 0,
            DestinationSize = File.Exists(destPath) ? new FileInfo(destPath).Length : 0,
            SourceDate = File.GetLastWriteTime(sourcePath),
            DestinationDate = File.Exists(destPath) ? File.GetLastWriteTime(destPath) : default
        };

        var resolution = await ConflictResolutionRequested(conflict);

        return resolution.Resolution switch
        {
            ConflictResolution.Replace => destPath,
            ConflictResolution.Skip => null,
            ConflictResolution.RenameWithSuffix => GenerateAutoSuffixPath(destPath),
            ConflictResolution.Cancel => throw new OperationCanceledException("User cancelled operation"),
            _ => destPath
        };
    }

    private static string GenerateAutoSuffixPath(string path)
    {
        var dir = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        int counter = 2;
        string newPath;
        do
        {
            newPath = Path.Combine(dir, $"{name} ({counter++}){ext}");
        } while (File.Exists(newPath) || Directory.Exists(newPath));
        return newPath;
    }

    private static string GenerateExplorerCopyPath(string existingPath)
    {
        var dir = Path.GetDirectoryName(existingPath) ?? string.Empty;
        var isDirectory = Directory.Exists(existingPath);
        var ext = isDirectory ? string.Empty : Path.GetExtension(existingPath);
        var name = isDirectory
            ? Path.GetFileName(existingPath)
            : Path.GetFileNameWithoutExtension(existingPath);

        var candidate = Path.Combine(dir, $"{name} - Copy{ext}");
        if (!File.Exists(candidate) && !Directory.Exists(candidate))
            return candidate;

        var counter = 2;
        while (true)
        {
            candidate = Path.Combine(dir, $"{name} - Copy ({counter}){ext}");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
                return candidate;

            counter++;
        }
    }

    private static string NormalizeDirectoryPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private static void CopyDirectoryRecursive(string sourceDir, string destDir, CancellationToken ct)
    {
        Directory.CreateDirectory(destDir);
        foreach (var file in Directory.EnumerateFiles(sourceDir))
        {
            ct.ThrowIfCancellationRequested();
            File.Copy(file, Path.Combine(destDir, Path.GetFileName(file)), overwrite: true);
        }
        foreach (var dir in Directory.EnumerateDirectories(sourceDir))
        {
            ct.ThrowIfCancellationRequested();
            CopyDirectoryRecursive(dir, Path.Combine(destDir, Path.GetFileName(dir)), ct);
        }
    }

    private void RaiseOperationError(string path, string message, int win32Error, bool isAccessDenied)
    {
        OperationError?.Invoke(this, new FileOperationErrorEventArgs
        {
            FilePath = path,
            ErrorMessage = message,
            Win32ErrorCode = win32Error,
            IsAccessDenied = isAccessDenied
        });
    }

    private static DateTimeOffset FileTimeToDateTimeOffset(System.Runtime.InteropServices.ComTypes.FILETIME fileTime)
    {
        long ticks = ((long)(uint)fileTime.dwHighDateTime << 32) | (uint)fileTime.dwLowDateTime;
        if (ticks <= 0)
            return DateTimeOffset.MinValue;

        return DateTimeOffset.FromFileTime(ticks);
    }

    #endregion

    #region Disposal

    /// <inheritdoc/>
    public void Dispose()
    {
        _serviceCts.Cancel();
        _operationChannel.Writer.TryComplete();
        try { _processorTask.Wait(TimeSpan.FromSeconds(5)); } catch { /* shutdown */ }
        _serviceCts.Dispose();
    }

    #endregion

    #region Inner Types

    private sealed record QueuedOperation(FileOperation Operation, IProgress<FileOperationProgress>? Progress);

    private sealed class WatcherHandle : IDisposable
    {
        private readonly FileSystemWatcher _watcher;
        private readonly ChangeDebouncer _debouncer;

        public WatcherHandle(FileSystemWatcher watcher, ChangeDebouncer debouncer)
        {
            _watcher = watcher;
            _debouncer = debouncer;
        }

        public void Dispose()
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Dispose();
            _debouncer.Dispose();
        }
    }

    private sealed class ChangeDebouncer : IDisposable
    {
        private readonly Action<FileSystemChangeEvent> _callback;
        private readonly TimeSpan _delay;
        private readonly ConcurrentQueue<FileSystemChangeEvent> _queue = new();
        private Timer? _timer;
        private int _processing;

        public ChangeDebouncer(Action<FileSystemChangeEvent> callback, TimeSpan delay)
        {
            _callback = callback;
            _delay = delay;
        }

        public void Enqueue(FileSystemChangeEvent change)
        {
            _queue.Enqueue(change);
            _timer?.Dispose();
            _timer = new Timer(FlushCallback, null, _delay, Timeout.InfiniteTimeSpan);
        }

        private void FlushCallback(object? state)
        {
            if (Interlocked.CompareExchange(ref _processing, 1, 0) != 0) return;
            try
            {
                while (_queue.TryDequeue(out var change))
                {
                    _callback(change);
                }
            }
            finally
            {
                Interlocked.Exchange(ref _processing, 0);
            }
        }

        public void Dispose() => _timer?.Dispose();
    }

    #endregion
}
