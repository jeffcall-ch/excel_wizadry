// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents a navigation history entry for tab navigation (Back/Forward).
/// </summary>
public sealed class DirectoryEntry
{
    /// <summary>Full path of the directory.</summary>
    public required string Path { get; init; }

    /// <summary>Display name for the directory.</summary>
    public required string DisplayName { get; init; }

    /// <summary>Timestamp when this entry was visited.</summary>
    public DateTimeOffset VisitedAt { get; init; } = DateTimeOffset.Now;
}

/// <summary>
/// Navigation history stack with Back/Forward support per tab.
/// </summary>
public sealed class NavigationHistory
{
    private readonly List<DirectoryEntry> _entries = [];
    private int _currentIndex = -1;

    /// <summary>Whether the user can navigate back.</summary>
    public bool CanGoBack => _currentIndex > 0;

    /// <summary>Whether the user can navigate forward.</summary>
    public bool CanGoForward => _currentIndex < _entries.Count - 1;

    /// <summary>Current directory entry, if any.</summary>
    public DirectoryEntry? Current => _currentIndex >= 0 && _currentIndex < _entries.Count
        ? _entries[_currentIndex]
        : null;

    /// <summary>Navigate to a new directory, clearing forward history.</summary>
    public void Navigate(DirectoryEntry entry)
    {
        if (_currentIndex < _entries.Count - 1)
        {
            _entries.RemoveRange(_currentIndex + 1, _entries.Count - _currentIndex - 1);
        }
        _entries.Add(entry);
        _currentIndex = _entries.Count - 1;
    }

    /// <summary>Go back one entry. Returns the entry navigated to.</summary>
    public DirectoryEntry? GoBack()
    {
        if (!CanGoBack) return null;
        _currentIndex--;
        return _entries[_currentIndex];
    }

    /// <summary>Go forward one entry. Returns the entry navigated to.</summary>
    public DirectoryEntry? GoForward()
    {
        if (!CanGoForward) return null;
        _currentIndex++;
        return _entries[_currentIndex];
    }

    /// <summary>Get all back entries for display in a dropdown.</summary>
    public IReadOnlyList<DirectoryEntry> GetBackEntries()
    {
        if (_currentIndex <= 0) return [];
        return _entries.GetRange(0, _currentIndex).AsReadOnly();
    }

    /// <summary>Get all forward entries for display in a dropdown.</summary>
    public IReadOnlyList<DirectoryEntry> GetForwardEntries()
    {
        if (_currentIndex >= _entries.Count - 1) return [];
        return _entries.GetRange(_currentIndex + 1, _entries.Count - _currentIndex - 1).AsReadOnly();
    }
}
