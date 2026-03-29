// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Core.ViewModels;

/// <summary>
/// ViewModel for the tree view (left navigation pane) showing Quick Access, This PC, and Network.
/// Supports lazy-loading of tree node children and synchronization with the active navigation path.
/// </summary>
public sealed partial class TreeViewModel : ObservableObject, IDisposable
{
    private readonly IFileSystemService _fileSystemService;
    private readonly IShellIntegrationService _shellService;
    private readonly ILogger<TreeViewModel> _logger;

    /// <summary>Root-level tree nodes.</summary>
    public ObservableCollection<TreeNodeViewModel> RootNodes { get; } = [];

    [ObservableProperty]
    private TreeNodeViewModel? _selectedNode;

    /// <summary>Raised when the user selects a tree node to navigate to that path.</summary>
    public event EventHandler<string>? NavigationRequested;

    /// <summary>Initializes a new instance of the <see cref="TreeViewModel"/> class.</summary>
    public TreeViewModel(
        IFileSystemService fileSystemService,
        IShellIntegrationService shellService,
        ILogger<TreeViewModel> logger)
    {
        _fileSystemService = fileSystemService;
        _shellService = shellService;
        _logger = logger;
    }

    /// <summary>Loads the initial tree structure.</summary>
    public async Task LoadAsync(CancellationToken cancellationToken = default)
    {
        RootNodes.Clear();

        // Quick Access
        var quickAccess = new TreeNodeViewModel(new TreeNode
        {
            Path = "shell:Quick Access",
            DisplayName = "Quick Access",
            NodeType = TreeNodeType.QuickAccess,
            IconGlyph = "\uE728",
            HasChildren = true
        });
        await LoadQuickAccessChildrenAsync(quickAccess, cancellationToken);
        RootNodes.Add(quickAccess);

        // This PC
        var thisPC = new TreeNodeViewModel(new TreeNode
        {
            Path = "shell:ThisPC",
            DisplayName = "This PC",
            NodeType = TreeNodeType.ThisPC,
            IconGlyph = "\uE7F8",
            HasChildren = true
        });
        await LoadDrivesAsync(thisPC, cancellationToken);
        RootNodes.Add(thisPC);

        // Network
        var network = new TreeNodeViewModel(new TreeNode
        {
            Path = "shell:NetworkPlacesFolder",
            DisplayName = "Network",
            NodeType = TreeNodeType.Network,
            IconGlyph = "\uE968",
            HasChildren = true
        });
        RootNodes.Add(network);
    }

    /// <summary>Synchronizes the tree selection with the given filesystem path, expanding intermediate nodes as needed.</summary>
    public async Task<TreeNodeViewModel?> SyncWithPathAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;

        var normalizedPath = NormalizeFileSystemPath(path);
        SelectedNode = null;

        foreach (var root in RootNodes)
        {
            var found = await ExpandAndSelectAsync(root, normalizedPath);
            if (found) break;
        }

        return SelectedNode;
    }

    private async Task<bool> ExpandAndSelectAsync(TreeNodeViewModel node, string targetPath)
    {
        var nodePath = node.Model.Path;

        if (string.IsNullOrWhiteSpace(nodePath))
            return false;

        // Skip shell paths that can't be ancestors
        if (nodePath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            // Check pinned folders under Quick Access — they have real paths
            if (!node.IsLoaded && node.Model.HasChildren)
                await ExpandNodeAsync(node);

            foreach (var child in node.Children)
            {
                if (string.IsNullOrWhiteSpace(child.Model.Path))
                    continue;

                var found = await ExpandAndSelectAsync(child, targetPath);
                if (found)
                {
                    node.IsExpanded = true;
                    return true;
                }
            }
            return false;
        }

        // Normalize node path for comparison
        var normalizedNodePath = NormalizeFileSystemPath(nodePath);

        // Exact match — select this node
        if (normalizedNodePath.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
        {
            if (!node.IsLoaded && node.Model.HasChildren)
            {
                await ExpandNodeAsync(node);
            }

            SelectedNode = node;
            node.IsExpanded = true;
            return true;
        }

        // Check if target is under this node's path
        var prefix = HasTrailingDirectorySeparator(normalizedNodePath)
            ? normalizedNodePath
            : normalizedNodePath + Path.DirectorySeparatorChar;

        if (!targetPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            return false;

        // Target is a descendant — expand this node if needed
        if (!node.IsLoaded && node.Model.HasChildren)
        {
            await ExpandNodeAsync(node);
        }
        node.IsExpanded = true;

        // Search children
        foreach (var child in node.Children)
        {
            if (string.IsNullOrWhiteSpace(child.Model.Path))
                continue;

            var found = await ExpandAndSelectAsync(child, targetPath);
            if (found) return true;
        }

        // If exact child not found, select this node as the closest ancestor
        if (!node.IsLoaded && node.Model.HasChildren)
        {
            await ExpandNodeAsync(node);
        }

        SelectedNode = node;
        return true;
    }

    private static string NormalizeFileSystemPath(string path)
    {
        var fullPath = Path.GetFullPath(path);
        var root = Path.GetPathRoot(fullPath);

        if (!string.IsNullOrEmpty(root) && fullPath.Equals(root, StringComparison.OrdinalIgnoreCase))
            return root;

        return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private static bool HasTrailingDirectorySeparator(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;

        var lastChar = path[^1];
        return lastChar == Path.DirectorySeparatorChar || lastChar == Path.AltDirectorySeparatorChar;
    }

    /// <summary>Expands a tree node and loads its children lazily.</summary>
    [RelayCommand]
    public async Task ExpandNodeAsync(TreeNodeViewModel node)
    {
        if (node.IsLoaded) return;

        try
        {
            if (node.Model.NodeType == TreeNodeType.Drive || node.Model.NodeType == TreeNodeType.Folder || node.Model.NodeType == TreeNodeType.PinnedFolder)
            {
                await LoadFolderChildrenAsync(node);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to expand node {Path}", node.Model.Path);
        }
    }

    /// <summary>Handles tree node selection — navigates to the selected path.</summary>
    [RelayCommand]
    private void SelectNode(TreeNodeViewModel node)
    {
        SelectedNode = node;
        var path = node.Model.Path;

        // Resolve shell paths
        if (path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            var resolved = _shellService.ResolveShellPath(path);
            if (resolved is not null) path = resolved;
            else return; // Cannot resolve
        }

        NavigationRequested?.Invoke(this, path);
    }

    private async Task LoadQuickAccessChildrenAsync(TreeNodeViewModel parent, CancellationToken ct)
    {
        var knownFolders = new (Guid Id, string Name, string Glyph)[]
        {
            (new("B4BFCC3A-DB2C-424C-B029-7FE99A87C641"), "Desktop", "\uE8FC"),
            (new("374DE290-123F-4565-9164-39C4925E467B"), "Downloads", "\uE896"),
            (new("FDD39AD0-238F-46AF-ADB4-6C85480369C7"), "Documents", "\uE8A5"),
            (new("33E28130-4E1E-4676-835A-98395C3BC3BB"), "Pictures", "\uEB9F"),
            (new("4BD8D571-6D19-48D3-BE97-422220080E43"), "Music", "\uEC4F"),
            (new("18989B1D-99B5-455B-841C-AB7C74E4DDFC"), "Videos", "\uE714"),
        };

        foreach (var (id, name, glyph) in knownFolders)
        {
            ct.ThrowIfCancellationRequested();
            var path = _shellService.ResolveShellPath($"shell:{name}");
            if (path is null || !Directory.Exists(path)) continue;

            parent.Children.Add(new TreeNodeViewModel(new TreeNode
            {
                Path = path,
                DisplayName = name,
                NodeType = TreeNodeType.PinnedFolder,
                IconGlyph = glyph,
                HasChildren = true
            }));
        }

        parent.IsLoaded = true;
        await Task.CompletedTask;
    }

    private async Task LoadDrivesAsync(TreeNodeViewModel parent, CancellationToken ct)
    {
        await Task.Run(() =>
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var label = drive.IsReady ? drive.VolumeLabel : string.Empty;
                    var displayName = string.IsNullOrEmpty(label)
                        ? $"Local Disk ({drive.Name.TrimEnd('\\')})"
                        : $"{label} ({drive.Name.TrimEnd('\\')})";

                    string? freeText = null;
                    double? freePct = null;

                    if (drive.IsReady)
                    {
                        var freeGb = drive.AvailableFreeSpace / (1024.0 * 1024 * 1024);
                        var totalGb = drive.TotalSize / (1024.0 * 1024 * 1024);
                        freeText = $"{freeGb:F1} GB free of {totalGb:F1} GB";
                        freePct = totalGb > 0 ? (1 - freeGb / totalGb) * 100 : 0;
                    }

                    var glyph = drive.DriveType switch
                    {
                        DriveType.Fixed => "\uEDA2",
                        DriveType.Removable => "\uE88E",
                        DriveType.CDRom => "\uE958",
                        DriveType.Network => "\uE968",
                        _ => "\uEDA2"
                    };

                    parent.Children.Add(new TreeNodeViewModel(new TreeNode
                    {
                        Path = drive.Name,
                        DisplayName = displayName,
                        NodeType = TreeNodeType.Drive,
                        IconGlyph = glyph,
                        HasChildren = true,
                        FreeSpaceText = freeText,
                        FreeSpacePercent = freePct
                    }));
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Failed to enumerate drive {Drive}", drive.Name);
                }
            }
        }, ct);

        parent.IsLoaded = true;
    }

    private async Task LoadFolderChildrenAsync(TreeNodeViewModel parent)
    {
        parent.Children.Clear();
        var path = parent.Model.Path;

        await Task.Run(() =>
        {
            try
            {
                var dirInfo = new DirectoryInfo(path);
                foreach (var dir in dirInfo.EnumerateDirectories())
                {
                    if ((dir.Attributes & FileAttributes.Hidden) != 0 || (dir.Attributes & FileAttributes.System) != 0)
                        continue;

                    parent.Children.Add(new TreeNodeViewModel(new TreeNode
                    {
                        Path = dir.FullName,
                        DisplayName = dir.Name,
                        NodeType = TreeNodeType.Folder,
                        IconGlyph = "\uE8B7",
                        HasChildren = HasSubDirectories(dir)
                    }));
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Silently skip inaccessible folders
            }
        });

        parent.IsLoaded = true;
    }

    private static bool HasSubDirectories(DirectoryInfo dir)
    {
        try
        {
            return dir.EnumerateDirectories().Any();
        }
        catch
        {
            return false;
        }
    }

    private bool FindAndSelectNode(TreeNodeViewModel node, string targetPath)
    {
        if (node.Model.Path.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
        {
            SelectedNode = node;
            return true;
        }

        foreach (var child in node.Children)
        {
            if (FindAndSelectNode(child, targetPath)) return true;
        }

        return false;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        // No unmanaged resources
    }
}

/// <summary>
/// ViewModel wrapper for a <see cref="TreeNode"/> for UI binding and lazy-loading state.
/// </summary>
public sealed partial class TreeNodeViewModel : ObservableObject
{
    /// <summary>The underlying tree node model.</summary>
    public TreeNode Model { get; }

    /// <summary>Children of this node.</summary>
    public ObservableCollection<TreeNodeViewModel> Children { get; } = [];

    [ObservableProperty]
    private bool _isExpanded;

    [ObservableProperty]
    private bool _isSelected;

    /// <summary>Whether children have been loaded.</summary>
    public bool IsLoaded { get; set; }

    /// <summary>Display name for UI binding.</summary>
    public string DisplayName => Model.DisplayName;

    /// <summary>Icon glyph for UI binding.</summary>
    public string IconGlyph => Model.IconGlyph;

    /// <summary>Free space text for drive nodes.</summary>
    public string? FreeSpaceText => Model.FreeSpaceText;

    /// <summary>Whether this node has children (for expand arrow).</summary>
    public bool HasChildren => Model.HasChildren;

    /// <summary>Initializes a new instance of the <see cref="TreeNodeViewModel"/> class.</summary>
    public TreeNodeViewModel(TreeNode model)
    {
        Model = model;
        if (model.HasChildren && !model.IsLoaded)
        {
            // Add a placeholder for the expand arrow
            Children.Add(new TreeNodeViewModel(new TreeNode
            {
                Path = string.Empty,
                DisplayName = "Loading...",
                NodeType = TreeNodeType.Folder,
                HasChildren = false
            }));
        }
    }

    /// <summary>Returns the display name (used as fallback when DataTemplate is not compiled).</summary>
    public override string ToString() => DisplayName;
}
