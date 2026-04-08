// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Text.Json;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Settings service that persists application configuration to the local app data folder
/// using JSON serialization. Compatible with both packaged (MSIX) and unpackaged deployments.
/// </summary>
public sealed class SettingsService : ISettingsService
{
    private readonly ILogger<SettingsService> _logger;
    private readonly string _settingsDir;
    private readonly string _settingsFilePath;
    private readonly string _sessionFilePath;
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <inheritdoc/>
    public AppSettings Settings { get; private set; } = new();

    /// <summary>Initializes a new instance of the <see cref="SettingsService"/> class.</summary>
    public SettingsService(ILogger<SettingsService> logger)
    {
        _logger = logger;
        _settingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FileExplorer");
        _settingsFilePath = Path.Combine(_settingsDir, "settings.json");
        _sessionFilePath = Path.Combine(_settingsDir, "session.json");
    }

    /// <inheritdoc/>
    public async Task LoadAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            if (!File.Exists(_settingsFilePath))
            {
                _logger.LogInformation("No settings file found, using defaults");
                return;
            }

            var json = await File.ReadAllTextAsync(_settingsFilePath, cancellationToken);
            var loaded = JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions);
            if (loaded is not null)
            {
                NormalizePerFolderSettingsCollections(loaded);
                NormalizeFavoritesCollection(loaded);
                Settings = loaded;
                _logger.LogInformation("Settings loaded from {Path}", _settingsFilePath);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to load settings, using defaults");
            Settings = new AppSettings();
        }
    }

    /// <inheritdoc/>
    public async Task SaveAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            Directory.CreateDirectory(_settingsDir);
            var json = JsonSerializer.Serialize(Settings, _jsonOptions);
            await File.WriteAllTextAsync(_settingsFilePath, json, cancellationToken);
            _logger.LogDebug("Settings saved to {Path}", _settingsFilePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save settings");
        }
    }

    /// <inheritdoc/>
    public async Task SaveTabSessionAsync(IReadOnlyList<TabState> tabs, CancellationToken cancellationToken = default)
    {
        try
        {
            Directory.CreateDirectory(_settingsDir);
            var savedTabs = tabs.Select((t, i) => new SavedTabState
            {
                Path = t.CurrentPath,
                IsPinned = t.IsPinned,
                Index = i
            }).ToList();

            var json = JsonSerializer.Serialize(savedTabs, _jsonOptions);
            await File.WriteAllTextAsync(_sessionFilePath, json, cancellationToken);
            _logger.LogDebug("Tab session saved: {Count} tab(s)", savedTabs.Count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save tab session");
        }
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyList<SavedTabState>> LoadTabSessionAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            if (!File.Exists(_sessionFilePath)) return [];

            var json = await File.ReadAllTextAsync(_sessionFilePath, cancellationToken);
            var tabs = JsonSerializer.Deserialize<List<SavedTabState>>(json, _jsonOptions);
            if (tabs is not null)
            {
                _logger.LogInformation("Tab session loaded: {Count} tab(s)", tabs.Count);
                return tabs.OrderBy(t => t.Index).ToList();
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to load tab session");
        }
        return [];
    }

    /// <inheritdoc/>
    public void SaveColumnWidths(string folderPath, Dictionary<string, double> widths)
    {
        var normalizedKey = NormalizeFolderSettingsKey(folderPath);
        if (string.IsNullOrWhiteSpace(normalizedKey))
            return;

        Settings.PerFolderColumnWidths[normalizedKey] = widths;
    }

    /// <inheritdoc/>
    public Dictionary<string, double>? GetColumnWidths(string folderPath)
    {
        var normalizedKey = NormalizeFolderSettingsKey(folderPath);
        if (string.IsNullOrWhiteSpace(normalizedKey))
            return null;

        if (Settings.PerFolderColumnWidths.TryGetValue(normalizedKey, out var widths))
            return widths;

        var fallback = Settings.PerFolderColumnWidths
            .FirstOrDefault(pair => string.Equals(pair.Key, normalizedKey, StringComparison.OrdinalIgnoreCase));

        return fallback.Key is null ? null : fallback.Value;
    }

    /// <inheritdoc/>
    public void SaveFolderSort(string folderPath, FolderSortSettings sortSettings)
    {
        var normalizedKey = NormalizeFolderSettingsKey(folderPath);
        if (string.IsNullOrWhiteSpace(normalizedKey))
            return;

        Settings.PerFolderSortSettings[normalizedKey] = new FolderSortSettings
        {
            SortColumn = string.IsNullOrWhiteSpace(sortSettings.SortColumn) ? "Name" : sortSettings.SortColumn,
            SortAscending = sortSettings.SortAscending
        };
    }

    /// <inheritdoc/>
    public FolderSortSettings? GetFolderSort(string folderPath)
    {
        var normalizedKey = NormalizeFolderSettingsKey(folderPath);
        if (string.IsNullOrWhiteSpace(normalizedKey))
            return null;

        if (Settings.PerFolderSortSettings.TryGetValue(normalizedKey, out var sortSettings))
            return sortSettings;

        var fallback = Settings.PerFolderSortSettings
            .FirstOrDefault(pair => string.Equals(pair.Key, normalizedKey, StringComparison.OrdinalIgnoreCase));

        return fallback.Key is null ? null : fallback.Value;
    }

    private static void NormalizePerFolderSettingsCollections(AppSettings settings)
    {
        settings.PerFolderColumnWidths = settings.PerFolderColumnWidths
            .Where(pair => !string.IsNullOrWhiteSpace(pair.Key))
            .ToDictionary(
                pair => NormalizeFolderSettingsKey(pair.Key),
                pair => pair.Value,
                StringComparer.OrdinalIgnoreCase);

        settings.PerFolderSortSettings = settings.PerFolderSortSettings
            .Where(pair => !string.IsNullOrWhiteSpace(pair.Key))
            .ToDictionary(
                pair => NormalizeFolderSettingsKey(pair.Key),
                pair => pair.Value,
                StringComparer.OrdinalIgnoreCase);
    }

    private static void NormalizeFavoritesCollection(AppSettings settings)
    {
        var sourceFavorites = settings.Favorites;
        if (sourceFavorites.Count == 0 && settings.FavoritePaths.Count > 0)
        {
            sourceFavorites = settings.FavoritePaths
                .Select(path => new FavoriteEntrySettings { Path = path })
                .ToList();
        }

        var deduped = new Dictionary<string, FavoriteEntrySettings>(StringComparer.OrdinalIgnoreCase);
        foreach (var favorite in sourceFavorites)
        {
            var normalizedPath = NormalizeFolderSettingsKey(favorite.Path);
            if (string.IsNullOrWhiteSpace(normalizedPath))
                continue;

            var normalizedAlias = NormalizeFavoriteAlias(favorite.Alias, normalizedPath);

            if (!deduped.TryGetValue(normalizedPath, out var existing))
            {
                deduped[normalizedPath] = new FavoriteEntrySettings
                {
                    Path = normalizedPath,
                    Alias = normalizedAlias
                };

                continue;
            }

            if (string.IsNullOrWhiteSpace(existing.Alias) && !string.IsNullOrWhiteSpace(normalizedAlias))
            {
                existing.Alias = normalizedAlias;
            }
        }

        settings.Favorites = deduped.Values
            .OrderBy(GetFavoriteSortLabel, StringComparer.OrdinalIgnoreCase)
            .ThenBy(favorite => favorite.Path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Legacy property is retained only for migration.
        settings.FavoritePaths = [];
    }

    private static string GetFavoriteSortLabel(FavoriteEntrySettings favorite)
    {
        return string.IsNullOrWhiteSpace(favorite.Alias)
            ? GetPathLeafLabel(favorite.Path)
            : favorite.Alias;
    }

    private static string? NormalizeFavoriteAlias(string? alias, string path)
    {
        if (string.IsNullOrWhiteSpace(alias))
            return null;

        var trimmed = alias.Trim();
        return string.Equals(trimmed, GetPathLeafLabel(path), StringComparison.OrdinalIgnoreCase)
            ? null
            : trimmed;
    }

    private static string GetPathLeafLabel(string path)
    {
        var trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var leaf = Path.GetFileName(trimmed);
        return string.IsNullOrWhiteSpace(leaf) ? path : leaf;
    }

    private static string NormalizeFolderSettingsKey(string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
            return string.Empty;

        var candidate = StripLongPathPrefix(folderPath.Trim().Trim('"'))
            .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

        try
        {
            candidate = Path.GetFullPath(candidate);
        }
        catch
        {
            // Keep best-effort key when the path is transient or malformed.
        }

        var root = Path.GetPathRoot(candidate);
        if (!string.IsNullOrWhiteSpace(root) &&
            !candidate.Equals(root, StringComparison.OrdinalIgnoreCase))
        {
            candidate = candidate.TrimEnd(Path.DirectorySeparatorChar);
        }

        return candidate;
    }

    private static string StripLongPathPrefix(string path)
    {
        if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            return @"\\" + path[8..];

        if (path.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            return path[4..];

        return path;
    }
}
