// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Service for persisting and loading application settings.
/// </summary>
public interface ISettingsService
{
    /// <summary>Gets the current application settings.</summary>
    AppSettings Settings { get; }

    /// <summary>Loads settings from persistent storage.</summary>
    Task LoadAsync(CancellationToken cancellationToken = default);

    /// <summary>Saves settings to persistent storage.</summary>
    Task SaveAsync(CancellationToken cancellationToken = default);

    /// <summary>Saves the current tab session state.</summary>
    Task SaveTabSessionAsync(IReadOnlyList<TabState> tabs, CancellationToken cancellationToken = default);

    /// <summary>Loads the saved tab session state.</summary>
    Task<IReadOnlyList<SavedTabState>> LoadTabSessionAsync(CancellationToken cancellationToken = default);

    /// <summary>Saves column widths for a specific folder.</summary>
    void SaveColumnWidths(string folderPath, Dictionary<string, double> widths);

    /// <summary>Gets persisted column widths for a specific folder.</summary>
    Dictionary<string, double>? GetColumnWidths(string folderPath);

    /// <summary>Saves sort settings for a specific folder.</summary>
    void SaveFolderSort(string folderPath, FolderSortSettings sortSettings);

    /// <summary>Gets persisted sort settings for a specific folder.</summary>
    FolderSortSettings? GetFolderSort(string folderPath);
}
