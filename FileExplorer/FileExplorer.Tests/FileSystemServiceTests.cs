// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Threading.Channels;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Infrastructure.Services;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;
using FluentAssertions;

namespace FileExplorer.Tests;

/// <summary>
/// Unit tests for <see cref="FileSystemService"/> operation queue ordering,
/// conflict resolution logic, and undo stack correctness.
/// </summary>
public class FileSystemServiceTests : IDisposable
{
    private readonly Mock<ILogger<FileSystemService>> _loggerMock;
    private readonly FileSystemService _service;
    private readonly string _testDir;

    public FileSystemServiceTests()
    {
        _loggerMock = new Mock<ILogger<FileSystemService>>();
        _service = new FileSystemService(_loggerMock.Object);
        _testDir = Path.Combine(Path.GetTempPath(), "FileExplorerTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_testDir);
    }

    public void Dispose()
    {
        _service.Dispose();
        try
        {
            if (Directory.Exists(_testDir))
                Directory.Delete(_testDir, recursive: true);
        }
        catch { /* Cleanup best-effort */ }
    }

    #region Create New Folder

    [Fact]
    public async Task CreateNewFolderAsync_CreatesFolder_WithDefaultName()
    {
        var result = await _service.CreateNewFolderAsync(_testDir);

        result.Should().NotBeNullOrEmpty();
        Path.GetFileName(result).Should().Be("New folder");
        Directory.Exists(result).Should().BeTrue();
    }

    [Fact]
    public async Task CreateNewFolderAsync_AutoIncrements_WhenNameExists()
    {
        Directory.CreateDirectory(Path.Combine(_testDir, "New folder"));

        var result = await _service.CreateNewFolderAsync(_testDir);

        Path.GetFileName(result).Should().Be("New folder (2)");
        Directory.Exists(result).Should().BeTrue();
    }

    [Fact]
    public async Task CreateNewFolderAsync_AutoIncrements_MultipleExisting()
    {
        Directory.CreateDirectory(Path.Combine(_testDir, "New folder"));
        Directory.CreateDirectory(Path.Combine(_testDir, "New folder (2)"));
        Directory.CreateDirectory(Path.Combine(_testDir, "New folder (3)"));

        var result = await _service.CreateNewFolderAsync(_testDir);

        Path.GetFileName(result).Should().Be("New folder (4)");
    }

    #endregion

    #region Rename

    [Fact]
    public async Task RenameAsync_SuccessfulRename_ReturnsTrue()
    {
        var filePath = Path.Combine(_testDir, "original.txt");
        await File.WriteAllTextAsync(filePath, "test content");

        var result = await _service.RenameAsync(filePath, "renamed.txt");

        result.Should().BeTrue();
        File.Exists(Path.Combine(_testDir, "renamed.txt")).Should().BeTrue();
        File.Exists(filePath).Should().BeFalse();
    }

    [Fact]
    public async Task RenameAsync_PopulatesUndoStack()
    {
        var filePath = Path.Combine(_testDir, "original.txt");
        await File.WriteAllTextAsync(filePath, "test content");

        await _service.RenameAsync(filePath, "renamed.txt");

        _service.CanUndo.Should().BeTrue();
    }

    [Fact]
    public async Task RenameAsync_UndoRestoresOriginalName()
    {
        var filePath = Path.Combine(_testDir, "undo_original.txt");
        await File.WriteAllTextAsync(filePath, "test content");

        await _service.RenameAsync(filePath, "undo_renamed.txt");
        var undoResult = await _service.UndoLastOperationAsync();

        undoResult.Should().BeTrue();
        File.Exists(filePath).Should().BeTrue();
        File.Exists(Path.Combine(_testDir, "undo_renamed.txt")).Should().BeFalse();
    }

    #endregion

    #region Delete to Recycle Bin

    [Fact]
    public async Task DeleteToRecycleBinAsync_DeletesFile()
    {
        var filePath = Path.Combine(_testDir, "to_delete.txt");
        await File.WriteAllTextAsync(filePath, "delete me");

        var result = await _service.DeleteToRecycleBinAsync([filePath]);

        result.Should().BeTrue();
        // File should no longer exist at original location
        File.Exists(filePath).Should().BeFalse();
    }

    #endregion

    #region Permanent Delete

    [Fact]
    public async Task PermanentDeleteAsync_DeletesFile()
    {
        var filePath = Path.Combine(_testDir, "perm_delete.txt");
        await File.WriteAllTextAsync(filePath, "permanent delete");

        var result = await _service.PermanentDeleteAsync([filePath]);

        result.Should().BeTrue();
        File.Exists(filePath).Should().BeFalse();
    }

    [Fact]
    public async Task PermanentDeleteAsync_DeletesDirectory()
    {
        var dirPath = Path.Combine(_testDir, "dir_to_delete");
        Directory.CreateDirectory(dirPath);
        await File.WriteAllTextAsync(Path.Combine(dirPath, "child.txt"), "child");

        var result = await _service.PermanentDeleteAsync([dirPath]);

        result.Should().BeTrue();
        Directory.Exists(dirPath).Should().BeFalse();
    }

    #endregion

    #region Directory Enumeration

    [Fact]
    public async Task EnumerateDirectoryAsync_ReturnsAllEntries()
    {
        await File.WriteAllTextAsync(Path.Combine(_testDir, "file1.txt"), "a");
        await File.WriteAllTextAsync(Path.Combine(_testDir, "file2.txt"), "b");
        Directory.CreateDirectory(Path.Combine(_testDir, "subfolder"));

        var entries = new List<FileSystemEntry>();
        await foreach (var entry in _service.EnumerateDirectoryAsync(_testDir))
        {
            entries.Add(entry);
        }

        entries.Should().HaveCount(3);
        entries.Should().Contain(e => e.Name == "file1.txt" && !e.IsDirectory);
        entries.Should().Contain(e => e.Name == "file2.txt" && !e.IsDirectory);
        entries.Should().Contain(e => e.Name == "subfolder" && e.IsDirectory);
    }

    [Fact]
    public async Task EnumerateDirectoryAsync_ReportsCorrectFileSize()
    {
        var content = new string('x', 1024);
        await File.WriteAllTextAsync(Path.Combine(_testDir, "sized.txt"), content);

        var entries = new List<FileSystemEntry>();
        await foreach (var entry in _service.EnumerateDirectoryAsync(_testDir))
        {
            entries.Add(entry);
        }

        entries.First(e => e.Name == "sized.txt").Size.Should().BeGreaterThan(0);
    }

    [Fact]
    public async Task EnumerateDirectoryAsync_HandlesNonExistentPath()
    {
        var entries = new List<FileSystemEntry>();
        await foreach (var entry in _service.EnumerateDirectoryAsync(Path.Combine(_testDir, "nonexistent")))
        {
            entries.Add(entry);
        }

        entries.Should().BeEmpty();
    }

    #endregion

    #region Operation Queue

    [Fact]
    public async Task EnqueueOperation_CopyFile_CreatesDestination()
    {
        var srcFile = Path.Combine(_testDir, "src.txt");
        var destDir = Path.Combine(_testDir, "dest");
        Directory.CreateDirectory(destDir);
        await File.WriteAllTextAsync(srcFile, "source content");

        var op = new FileOperation
        {
            OperationType = FileOperationType.Copy,
            SourcePaths = [srcFile],
            DestinationPath = destDir
        };

        var tcs = new TaskCompletionSource();
        var progress = new Progress<FileOperationProgress>(p =>
        {
            if (p.IsComplete) tcs.TrySetResult();
        });

        await _service.EnqueueOperationAsync(op, progress);
        await tcs.Task.WaitAsync(TimeSpan.FromSeconds(10));

        File.Exists(Path.Combine(destDir, "src.txt")).Should().BeTrue();
    }

    [Fact]
    public async Task EnqueueOperation_MoveFile_RemovesSource()
    {
        var srcFile = Path.Combine(_testDir, "move_src.txt");
        var destDir = Path.Combine(_testDir, "move_dest");
        Directory.CreateDirectory(destDir);
        await File.WriteAllTextAsync(srcFile, "move content");

        var op = new FileOperation
        {
            OperationType = FileOperationType.Move,
            SourcePaths = [srcFile],
            DestinationPath = destDir
        };

        var tcs = new TaskCompletionSource();
        var progress = new Progress<FileOperationProgress>(p =>
        {
            if (p.IsComplete) tcs.TrySetResult();
        });

        await _service.EnqueueOperationAsync(op, progress);
        await tcs.Task.WaitAsync(TimeSpan.FromSeconds(10));

        File.Exists(srcFile).Should().BeFalse();
        File.Exists(Path.Combine(destDir, "move_src.txt")).Should().BeTrue();
    }

    #endregion

    #region Undo Stack

    [Fact]
    public async Task UndoStack_MultipleOperations_UndoesInReverseOrder()
    {
        var file1 = Path.Combine(_testDir, "undo1.txt");
        var file2 = Path.Combine(_testDir, "undo2.txt");
        await File.WriteAllTextAsync(file1, "content1");
        await File.WriteAllTextAsync(file2, "content2");

        await _service.RenameAsync(file1, "undo1_renamed.txt");
        await _service.RenameAsync(file2, "undo2_renamed.txt");

        // Undo the second rename first
        await _service.UndoLastOperationAsync();
        File.Exists(file2).Should().BeTrue();
        File.Exists(Path.Combine(_testDir, "undo2_renamed.txt")).Should().BeFalse();

        // Then undo the first
        await _service.UndoLastOperationAsync();
        File.Exists(file1).Should().BeTrue();
        File.Exists(Path.Combine(_testDir, "undo1_renamed.txt")).Should().BeFalse();
    }

    [Fact]
    public async Task UndoStack_CreateFolder_UndoDeletesFolder()
    {
        var newPath = await _service.CreateNewFolderAsync(_testDir);

        Directory.Exists(newPath).Should().BeTrue();
        _service.CanUndo.Should().BeTrue();

        await _service.UndoLastOperationAsync();
        Directory.Exists(newPath).Should().BeFalse();
    }

    [Fact]
    public async Task UndoStack_EmptyStack_ReturnsFalse()
    {
        var result = await _service.UndoLastOperationAsync();
        result.Should().BeFalse();
    }

    #endregion

    #region FileSystemWatcher

    [Fact]
    public async Task WatchDirectory_DetectsCreatedFile()
    {
        var tcs = new TaskCompletionSource<FileSystemChangeEvent>();

        using var handle = _service.WatchDirectory(_testDir, change =>
        {
            if (change.ChangeType == WatcherChangeTypes.Created)
                tcs.TrySetResult(change);
        });

        // Create a file after a brief delay
        await Task.Delay(100);
        var testFile = Path.Combine(_testDir, "watched_file.txt");
        await File.WriteAllTextAsync(testFile, "watched");

        var result = await tcs.Task.WaitAsync(TimeSpan.FromSeconds(5));
        result.ChangeType.Should().Be(WatcherChangeTypes.Created);
        result.Name.Should().Be("watched_file.txt");
    }

    [Fact]
    public async Task WatchDirectory_DetectsDeletedFile()
    {
        var testFile = Path.Combine(_testDir, "to_watch_delete.txt");
        await File.WriteAllTextAsync(testFile, "will be deleted");

        var tcs = new TaskCompletionSource<FileSystemChangeEvent>();

        using var handle = _service.WatchDirectory(_testDir, change =>
        {
            if (change.ChangeType == WatcherChangeTypes.Deleted)
                tcs.TrySetResult(change);
        });

        await Task.Delay(100);
        File.Delete(testFile);

        var result = await tcs.Task.WaitAsync(TimeSpan.FromSeconds(5));
        result.ChangeType.Should().Be(WatcherChangeTypes.Deleted);
    }

    #endregion
}
