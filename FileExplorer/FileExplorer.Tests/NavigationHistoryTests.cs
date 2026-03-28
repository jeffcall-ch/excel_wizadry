// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;
using Xunit;
using FluentAssertions;

namespace FileExplorer.Tests;

/// <summary>
/// Unit tests for <see cref="NavigationHistory"/> back/forward navigation logic.
/// </summary>
public class NavigationHistoryTests
{
    [Fact]
    public void Navigate_AddsEntry_UpdatesCurrent()
    {
        var history = new NavigationHistory();
        var entry = new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" };

        history.Navigate(entry);

        history.Current.Should().Be(entry);
    }

    [Fact]
    public void CanGoBack_IsFalse_WhenOnlyOneEntry()
    {
        var history = new NavigationHistory();
        history.Navigate(new DirectoryEntry { Path = @"C:\", DisplayName = "C:" });

        history.CanGoBack.Should().BeFalse();
    }

    [Fact]
    public void CanGoBack_IsTrue_WhenMultipleEntries()
    {
        var history = new NavigationHistory();
        history.Navigate(new DirectoryEntry { Path = @"C:\", DisplayName = "C:" });
        history.Navigate(new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" });

        history.CanGoBack.Should().BeTrue();
    }

    [Fact]
    public void GoBack_ReturnsPreviousEntry()
    {
        var history = new NavigationHistory();
        var first = new DirectoryEntry { Path = @"C:\", DisplayName = "C:" };
        var second = new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" };

        history.Navigate(first);
        history.Navigate(second);

        var result = history.GoBack();

        result.Should().Be(first);
        history.Current.Should().Be(first);
    }

    [Fact]
    public void GoForward_ReturnsNull_WhenNoForwardHistory()
    {
        var history = new NavigationHistory();
        history.Navigate(new DirectoryEntry { Path = @"C:\", DisplayName = "C:" });

        history.GoForward().Should().BeNull();
    }

    [Fact]
    public void GoForward_ReturnsNextEntry_AfterGoBack()
    {
        var history = new NavigationHistory();
        var first = new DirectoryEntry { Path = @"C:\", DisplayName = "C:" };
        var second = new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" };

        history.Navigate(first);
        history.Navigate(second);
        history.GoBack();

        var result = history.GoForward();

        result.Should().Be(second);
    }

    [Fact]
    public void Navigate_AfterGoBack_ClearsForwardHistory()
    {
        var history = new NavigationHistory();
        history.Navigate(new DirectoryEntry { Path = @"C:\", DisplayName = "C:" });
        history.Navigate(new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" });
        history.Navigate(new DirectoryEntry { Path = @"C:\Users\Public", DisplayName = "Public" });

        history.GoBack(); // at Users
        history.Navigate(new DirectoryEntry { Path = @"C:\Windows", DisplayName = "Windows" });

        history.CanGoForward.Should().BeFalse();
        history.Current!.Path.Should().Be(@"C:\Windows");
    }

    [Fact]
    public void GetBackEntries_ReturnsAllPreviousEntries()
    {
        var history = new NavigationHistory();
        var first = new DirectoryEntry { Path = @"C:\", DisplayName = "C:" };
        var second = new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" };
        var third = new DirectoryEntry { Path = @"C:\Users\Public", DisplayName = "Public" };

        history.Navigate(first);
        history.Navigate(second);
        history.Navigate(third);

        var backEntries = history.GetBackEntries();

        backEntries.Should().HaveCount(2);
        backEntries[0].Should().Be(first);
        backEntries[1].Should().Be(second);
    }

    [Fact]
    public void GetForwardEntries_ReturnsAllForwardEntries()
    {
        var history = new NavigationHistory();
        history.Navigate(new DirectoryEntry { Path = @"C:\", DisplayName = "C:" });
        history.Navigate(new DirectoryEntry { Path = @"C:\Users", DisplayName = "Users" });
        history.Navigate(new DirectoryEntry { Path = @"C:\Users\Public", DisplayName = "Public" });

        history.GoBack();
        history.GoBack();

        var forwardEntries = history.GetForwardEntries();

        forwardEntries.Should().HaveCount(2);
    }

    [Fact]
    public void GoBack_WhenCannotGoBack_ReturnsNull()
    {
        var history = new NavigationHistory();
        history.GoBack().Should().BeNull();
    }

    [Fact]
    public void Navigate_Multiple_TracksCorrectIndex()
    {
        var history = new NavigationHistory();
        for (int i = 0; i < 10; i++)
        {
            history.Navigate(new DirectoryEntry { Path = $@"C:\Folder{i}", DisplayName = $"Folder{i}" });
        }

        history.Current!.Path.Should().Be(@"C:\Folder9");
        history.CanGoBack.Should().BeTrue();
        history.CanGoForward.Should().BeFalse();
    }
}
