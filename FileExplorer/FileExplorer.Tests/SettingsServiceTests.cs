// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;
using FileExplorer.Infrastructure.Services;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;
using FluentAssertions;

namespace FileExplorer.Tests;

/// <summary>
/// Unit tests for <see cref="SettingsService"/> persistence and tab session management.
/// </summary>
public class SettingsServiceTests : IDisposable
{
    private readonly Mock<ILogger<SettingsService>> _loggerMock;
    private readonly SettingsService _service;
    private readonly string _tempSettingsDir;

    public SettingsServiceTests()
    {
        _loggerMock = new Mock<ILogger<SettingsService>>();
        _service = new SettingsService(_loggerMock.Object);

        _tempSettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FileExplorer");
    }

    public void Dispose()
    {
        // Cleanup is skipped for safety — settings directory is shared
    }

    [Fact]
    public void Settings_DefaultValues_AreCorrect()
    {
        _service.Settings.ShowHiddenFiles.Should().BeFalse();
        _service.Settings.ShowFileExtensions.Should().BeTrue();
        _service.Settings.ShowPreviewPane.Should().BeFalse();
        _service.Settings.NavigationPaneWidth.Should().Be(250);
        _service.Settings.RestoreSessionOnLaunch.Should().BeTrue();
    }

    [Fact]
    public async Task SaveAndLoad_RoundTrips_Settings()
    {
        _service.Settings.ShowHiddenFiles = true;
        _service.Settings.ShowPreviewPane = true;
        _service.Settings.NavigationPaneWidth = 350;

        await _service.SaveAsync();

        // Create new service instance to verify persistence
        var service2 = new SettingsService(_loggerMock.Object);
        await service2.LoadAsync();

        service2.Settings.ShowHiddenFiles.Should().BeTrue();
        service2.Settings.ShowPreviewPane.Should().BeTrue();
        service2.Settings.NavigationPaneWidth.Should().Be(350);
    }

    [Fact]
    public async Task SaveTabSession_And_LoadTabSession_RoundTrips()
    {
        var tabs = new List<TabState>
        {
            new() { CurrentPath = @"C:\Users", IsPinned = false },
            new() { CurrentPath = @"C:\Windows", IsPinned = true }
        };

        await _service.SaveTabSessionAsync(tabs);
        var loaded = await _service.LoadTabSessionAsync();

        loaded.Should().HaveCount(2);
        loaded[0].Path.Should().Be(@"C:\Users");
        loaded[0].IsPinned.Should().BeFalse();
        loaded[1].Path.Should().Be(@"C:\Windows");
        loaded[1].IsPinned.Should().BeTrue();
    }

    [Fact]
    public async Task LoadTabSession_ReturnsEmpty_WhenNoSessionFile()
    {
        // Delete session file if it exists
        var sessionPath = Path.Combine(_tempSettingsDir, "session.json");
        if (File.Exists(sessionPath)) File.Delete(sessionPath);

        var loaded = await _service.LoadTabSessionAsync();
        loaded.Should().BeEmpty();
    }

    [Fact]
    public void SaveColumnWidths_And_GetColumnWidths_RoundTrips()
    {
        var widths = new Dictionary<string, double>
        {
            ["Name"] = 300,
            ["DateModified"] = 180,
            ["Size"] = 100
        };

        _service.SaveColumnWidths(@"C:\Users", widths);
        var retrieved = _service.GetColumnWidths(@"C:\Users");

        retrieved.Should().NotBeNull();
        retrieved!["Name"].Should().Be(300);
        retrieved["DateModified"].Should().Be(180);
        retrieved["Size"].Should().Be(100);
    }

    [Fact]
    public void GetColumnWidths_ReturnsNull_ForUnknownPath()
    {
        var result = _service.GetColumnWidths(@"C:\UnknownPath\That\DoesNot\Exist");
        result.Should().BeNull();
    }

    [Fact]
    public void SaveFolderSort_And_GetFolderSort_RoundTrips()
    {
        _service.SaveFolderSort(@"C:\Users\szil\Downloads", new FolderSortSettings
        {
            SortColumn = "DateModified",
            SortAscending = false
        });

        var retrieved = _service.GetFolderSort(@"C:\Users\szil\Downloads");

        retrieved.Should().NotBeNull();
        retrieved!.SortColumn.Should().Be("DateModified");
        retrieved.SortAscending.Should().BeFalse();
    }

    [Fact]
    public void GetFolderSort_IsCaseInsensitive_AndPathNormalized()
    {
        _service.SaveFolderSort(@"C:\Users\szil\Downloads", new FolderSortSettings
        {
            SortColumn = "Size",
            SortAscending = true
        });

        var retrieved = _service.GetFolderSort(@"c:\USERS\szil\Downloads\\");

        retrieved.Should().NotBeNull();
        retrieved!.SortColumn.Should().Be("Size");
        retrieved.SortAscending.Should().BeTrue();
    }
}
