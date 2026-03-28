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
        Settings.PerFolderColumnWidths[folderPath] = widths;
    }

    /// <inheritdoc/>
    public Dictionary<string, double>? GetColumnWidths(string folderPath)
    {
        return Settings.PerFolderColumnWidths.TryGetValue(folderPath, out var widths) ? widths : null;
    }
}
