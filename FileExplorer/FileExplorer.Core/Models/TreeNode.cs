// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents a node in the navigation tree (left pane).
/// </summary>
public sealed class TreeNode
{
    /// <summary>Full path this node represents.</summary>
    public required string Path { get; init; }

    /// <summary>Display label for the node.</summary>
    public required string DisplayName { get; init; }

    /// <summary>Type of the tree node.</summary>
    public required TreeNodeType NodeType { get; init; }

    /// <summary>Icon glyph from Segoe Fluent Icons.</summary>
    public string IconGlyph { get; init; } = "\uE8B7";

    /// <summary>Children of this node. Null means not yet loaded.</summary>
    public List<TreeNode>? Children { get; set; }

    /// <summary>Whether children have been loaded.</summary>
    public bool IsLoaded => Children is not null;

    /// <summary>Whether this node is currently expanded.</summary>
    public bool IsExpanded { get; set; }

    /// <summary>Whether this node is currently selected.</summary>
    public bool IsSelected { get; set; }

    /// <summary>Free space text for drive nodes (e.g., "128 GB free of 512 GB").</summary>
    public string? FreeSpaceText { get; init; }

    /// <summary>Free space percentage (0-100) for drive nodes.</summary>
    public double? FreeSpacePercent { get; init; }

    /// <summary>Whether this node has children (for showing expand arrow before loading).</summary>
    public bool HasChildren { get; init; } = true;
}

/// <summary>
/// Types of tree nodes in the navigation pane.
/// </summary>
public enum TreeNodeType
{
    QuickAccess,
    PinnedFolder,
    RecentFolder,
    ThisPC,
    Drive,
    Folder,
    Network,
    NetworkLocation
}
