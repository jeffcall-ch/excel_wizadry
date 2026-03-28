// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Text.Json.Serialization;

namespace FileExplorer.Core.Models;

/// <summary>
/// Application settings that are persisted across sessions.
/// </summary>
public sealed class AppSettings
{
    /// <summary>Whether to show hidden files.</summary>
    public bool ShowHiddenFiles { get; set; }

    /// <summary>Whether to show file extensions.</summary>
    public bool ShowFileExtensions { get; set; } = true;

    /// <summary>Whether the preview pane is visible.</summary>
    public bool ShowPreviewPane { get; set; }

    /// <summary>Width of the navigation (tree view) pane.</summary>
    public double NavigationPaneWidth { get; set; } = 250;

    /// <summary>Width of the preview pane.</summary>
    public double PreviewPaneWidth { get; set; } = 300;

    /// <summary>Default path for a new tab. Empty string means "This PC".</summary>
    public string NewTabDefaultPath { get; set; } = string.Empty;

    /// <summary>Whether to restore previous session tabs on launch.</summary>
    public bool RestoreSessionOnLaunch { get; set; } = false;

    /// <summary>Whether to prompt before restoring session.</summary>
    public bool PromptBeforeRestore { get; set; }

    /// <summary>Saved tab states for session restore.</summary>
    public List<SavedTabState> SavedTabs { get; set; } = [];

    /// <summary>Persisted column widths per folder path. Key is folder path, value maps column name to width.</summary>
    public Dictionary<string, Dictionary<string, double>> PerFolderColumnWidths { get; set; } = new();
}

/// <summary>
/// Minimal tab state for session persistence.
/// </summary>
public sealed class SavedTabState
{
    /// <summary>Path the tab was showing.</summary>
    [JsonPropertyName("path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Whether the tab was pinned.</summary>
    [JsonPropertyName("pinned")]
    public bool IsPinned { get; set; }

    /// <summary>Position index of the tab.</summary>
    [JsonPropertyName("index")]
    public int Index { get; set; }
}
