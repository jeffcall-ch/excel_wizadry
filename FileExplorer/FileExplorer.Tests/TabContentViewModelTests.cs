// Copyright (c) FileExplorer contributors. All rights reserved.

using FluentAssertions;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Core.ViewModels;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;
using Xunit;

namespace FileExplorer.Tests;

/// <summary>
/// Unit tests for <see cref="TabContentViewModel"/> navigation path handling.
/// </summary>
public class TabContentViewModelTests
{
    [Fact]
    public async Task NavigateAsync_BareDriveLetter_NormalizesToDriveRoot()
    {
        var driveRoot = Path.GetPathRoot(Environment.SystemDirectory)!;
        var bareDrive = driveRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        var fileSystem = new Mock<IFileSystemService>(MockBehavior.Strict);
        fileSystem
            .Setup(service => service.EnumerateDirectoryAsync(driveRoot, It.IsAny<CancellationToken>()))
            .Returns(EmptyEntries());
        fileSystem
            .Setup(service => service.WatchDirectory(driveRoot, It.IsAny<Action<FileSystemChangeEvent>>()))
            .Returns(Mock.Of<IDisposable>());

        var shell = new Mock<IShellIntegrationService>(MockBehavior.Strict);
        var cloudStatus = new Mock<ICloudStatusService>(MockBehavior.Strict);
        cloudStatus.SetupGet(service => service.IsSyncRootDetected).Returns(false);

        using var viewModel = new TabContentViewModel(
            fileSystem.Object,
            shell.Object,
            cloudStatus.Object,
            Mock.Of<IPreviewService>(),
            NullLogger<TabContentViewModel>.Instance);

        await viewModel.NavigateAsync(bareDrive);

        viewModel.CurrentPath.Should().Be(driveRoot);
        viewModel.TabState.CurrentPath.Should().Be(driveRoot);
        fileSystem.Verify(service => service.EnumerateDirectoryAsync(driveRoot, It.IsAny<CancellationToken>()), Times.Once);
        fileSystem.Verify(service => service.WatchDirectory(driveRoot, It.IsAny<Action<FileSystemChangeEvent>>()), Times.Once);
    }

    [Fact]
    public async Task NavigateAsync_LongPathPrefix_StripsPrefixFromCurrentPath()
    {
        var systemDirectory = Environment.SystemDirectory;
        var longPath = @"\\?\" + systemDirectory;

        var fileSystem = new Mock<IFileSystemService>(MockBehavior.Strict);
        fileSystem
            .Setup(service => service.EnumerateDirectoryAsync(systemDirectory, It.IsAny<CancellationToken>()))
            .Returns(EmptyEntries());
        fileSystem
            .Setup(service => service.WatchDirectory(systemDirectory, It.IsAny<Action<FileSystemChangeEvent>>()))
            .Returns(Mock.Of<IDisposable>());

        var shell = new Mock<IShellIntegrationService>(MockBehavior.Strict);
        var cloudStatus = new Mock<ICloudStatusService>(MockBehavior.Strict);
        cloudStatus.SetupGet(service => service.IsSyncRootDetected).Returns(false);

        using var viewModel = new TabContentViewModel(
            fileSystem.Object,
            shell.Object,
            cloudStatus.Object,
            Mock.Of<IPreviewService>(),
            NullLogger<TabContentViewModel>.Instance);

        await viewModel.NavigateAsync(longPath);

        viewModel.CurrentPath.Should().Be(systemDirectory);
        viewModel.TabState.CurrentPath.Should().Be(systemDirectory);
    }

    [Fact]
    public async Task CommitRenameAsync_UpdatesPath_ResortsAndRequestsReveal()
    {
        var fileSystem = new Mock<IFileSystemService>(MockBehavior.Strict);
        fileSystem
            .Setup(service => service.RenameAsync(@"C:\Temp\Zeta", "Alpha", It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var shell = new Mock<IShellIntegrationService>(MockBehavior.Strict);
        shell.Setup(service => service.GetTypeDescription(@"C:\Temp\Alpha")).Returns("File folder");
        shell.Setup(service => service.GetIconIndex(@"C:\Temp\Alpha", true)).Returns(0);

        var cloudStatus = new Mock<ICloudStatusService>(MockBehavior.Strict);
        using var viewModel = new TabContentViewModel(
            fileSystem.Object,
            shell.Object,
            cloudStatus.Object,
            Mock.Of<IPreviewService>(),
            NullLogger<TabContentViewModel>.Instance);

        var renamedItem = new FileItemViewModel(new FileSystemEntry
        {
            FullPath = @"C:\Temp\Zeta",
            Name = "Zeta",
            IsDirectory = true,
            DateModified = DateTimeOffset.Now,
            Attributes = FileAttributes.Directory,
            TypeDescription = "File folder"
        })
        {
            IsRenaming = true,
            EditingName = "Alpha"
        };

        var existingItem = new FileItemViewModel(new FileSystemEntry
        {
            FullPath = @"C:\Temp\Beta",
            Name = "Beta",
            IsDirectory = true,
            DateModified = DateTimeOffset.Now,
            Attributes = FileAttributes.Directory,
            TypeDescription = "File folder"
        });

        viewModel.Items.Add(renamedItem);
        viewModel.Items.Add(existingItem);

        FileItemViewModel? revealItem = null;
        viewModel.RevealItemRequested += (_, item) => revealItem = item;

        await viewModel.CommitRenameAsync(renamedItem);

        renamedItem.Entry.FullPath.Should().Be(@"C:\Temp\Alpha");
        renamedItem.Name.Should().Be("Alpha");
        viewModel.Items[0].Should().Be(renamedItem);
        revealItem.Should().Be(renamedItem);
    }

    private static async IAsyncEnumerable<FileSystemEntry> EmptyEntries()
    {
        await Task.CompletedTask;
        yield break;
    }
}