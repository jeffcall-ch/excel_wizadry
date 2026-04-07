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
    public void SetSort_DateModifiedAscending_MixesFoldersAndFilesByTimestamp()
    {
        using var viewModel = new TabContentViewModel(
            Mock.Of<IFileSystemService>(),
            Mock.Of<IShellIntegrationService>(),
            Mock.Of<ICloudStatusService>(),
            Mock.Of<IPreviewService>(),
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            Mock.Of<ISettingsService>(),
            NullLogger<TabContentViewModel>.Instance);

        var oldestFile = CreateItem("oldest.txt", isDirectory: false, new DateTimeOffset(2026, 1, 1, 8, 0, 0, TimeSpan.Zero));
        var middleFolder = CreateItem("MiddleFolder", isDirectory: true, new DateTimeOffset(2026, 1, 2, 8, 0, 0, TimeSpan.Zero));
        var newestFolder = CreateItem("NewestFolder", isDirectory: true, new DateTimeOffset(2026, 1, 3, 8, 0, 0, TimeSpan.Zero));

        viewModel.Items.Add(newestFolder);
        viewModel.Items.Add(oldestFile);
        viewModel.Items.Add(middleFolder);

        viewModel.SetSort("DateModified");

        viewModel.Items.Select(i => i.Name).Should().Equal("oldest.txt", "MiddleFolder", "NewestFolder");
    }

    [Fact]
    public void SetSort_DateModifiedDescending_MixesFoldersAndFilesByTimestamp()
    {
        using var viewModel = new TabContentViewModel(
            Mock.Of<IFileSystemService>(),
            Mock.Of<IShellIntegrationService>(),
            Mock.Of<ICloudStatusService>(),
            Mock.Of<IPreviewService>(),
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            Mock.Of<ISettingsService>(),
            NullLogger<TabContentViewModel>.Instance);

        var oldestFolder = CreateItem("OldFolder", isDirectory: true, new DateTimeOffset(2026, 1, 1, 8, 0, 0, TimeSpan.Zero));
        var middleFile = CreateItem("middle.txt", isDirectory: false, new DateTimeOffset(2026, 1, 2, 8, 0, 0, TimeSpan.Zero));
        var newestFile = CreateItem("newest.txt", isDirectory: false, new DateTimeOffset(2026, 1, 3, 8, 0, 0, TimeSpan.Zero));

        viewModel.Items.Add(oldestFolder);
        viewModel.Items.Add(newestFile);
        viewModel.Items.Add(middleFile);

        viewModel.SetSort("DateModified");
        viewModel.SetSort("DateModified");

        viewModel.Items.Select(i => i.Name).Should().Equal("newest.txt", "middle.txt", "OldFolder");
    }

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
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            Mock.Of<ISettingsService>(),
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
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            Mock.Of<ISettingsService>(),
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
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            Mock.Of<ISettingsService>(),
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

    [Fact]
    public async Task SetSort_PersistsSortForCurrentFolder()
    {
        var systemDirectory = Environment.SystemDirectory;

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

        var settings = new Mock<ISettingsService>(MockBehavior.Strict);
        settings.SetupGet(service => service.Settings).Returns(new AppSettings());
        settings.Setup(service => service.GetFolderSort(systemDirectory)).Returns((FolderSortSettings?)null);
        settings.Setup(service => service.SaveFolderSort(systemDirectory, It.IsAny<FolderSortSettings>()));
        settings.Setup(service => service.SaveAsync(It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        using var viewModel = new TabContentViewModel(
            fileSystem.Object,
            shell.Object,
            cloudStatus.Object,
            Mock.Of<IPreviewService>(),
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            settings.Object,
            NullLogger<TabContentViewModel>.Instance);

        await viewModel.NavigateAsync(systemDirectory);
        viewModel.SetSort("DateModified");

        settings.Verify(
            service => service.SaveFolderSort(
                systemDirectory,
                It.Is<FolderSortSettings>(value =>
                    value.SortColumn == "DateModified" &&
                    value.SortAscending)),
            Times.Once);
        settings.Verify(service => service.SaveAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task NavigateAsync_RestoresPersistedSortForFolder()
    {
        var systemDirectory = Environment.SystemDirectory;

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

        var settings = new Mock<ISettingsService>(MockBehavior.Strict);
        settings.SetupGet(service => service.Settings).Returns(new AppSettings());
        settings.Setup(service => service.GetFolderSort(systemDirectory)).Returns(new FolderSortSettings
        {
            SortColumn = "DateModified",
            SortAscending = false
        });

        using var viewModel = new TabContentViewModel(
            fileSystem.Object,
            shell.Object,
            cloudStatus.Object,
            Mock.Of<IPreviewService>(),
            Mock.Of<IIdentityExtractionService>(),
            Mock.Of<IIdentityCacheService>(),
            Mock.Of<ITabPathStateService>(),
            settings.Object,
            NullLogger<TabContentViewModel>.Instance);

        await viewModel.NavigateAsync(systemDirectory);

        viewModel.SortColumn.Should().Be("DateModified");
        viewModel.SortAscending.Should().BeFalse();
    }

    private static async IAsyncEnumerable<FileSystemEntry> EmptyEntries()
    {
        await Task.CompletedTask;
        yield break;
    }

    private static FileItemViewModel CreateItem(string name, bool isDirectory, DateTimeOffset modified)
    {
        var fullPath = isDirectory ? $@"C:\Temp\{name}" : $@"C:\Temp\{name}";
        return new FileItemViewModel(new FileSystemEntry
        {
            FullPath = fullPath,
            Name = name,
            IsDirectory = isDirectory,
            DateModified = modified,
            Attributes = isDirectory ? FileAttributes.Directory : FileAttributes.Normal,
            Extension = isDirectory ? string.Empty : Path.GetExtension(name),
            TypeDescription = isDirectory ? "File folder" : "Text Document"
        });
    }
}