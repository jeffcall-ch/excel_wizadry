// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents a tab in the file explorer with its own navigation scope.
/// </summary>
public sealed class TabState
{
    /// <summary>Unique identifier for this tab.</summary>
    public Guid Id { get; init; } = Guid.NewGuid();

    /// <summary>Current directory path displayed in this tab.</summary>
    public string CurrentPath { get; set; } = string.Empty;

    /// <summary>Display title for the tab header.</summary>
    public string Title { get; set; } = "New tab";

    /// <summary>Full path shown as the tab tooltip.</summary>
    public string ToolTip => CurrentPath;

    /// <summary>Whether this tab is pinned (prevents accidental close).</summary>
    public bool IsPinned { get; set; }

    /// <summary>Navigation history for this tab.</summary>
    public NavigationHistory History { get; } = new();

    /// <summary>Current sort column name.</summary>
    public string SortColumn { get; set; } = "Name";

    /// <summary>Current sort direction.</summary>
    public bool SortAscending { get; set; } = true;

    /// <summary>Persisted column widths per directory (key = column name).</summary>
    public Dictionary<string, double> ColumnWidths { get; set; } = new();
}
