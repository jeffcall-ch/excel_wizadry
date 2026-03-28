// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Core.ViewModels;

/// <summary>
/// ViewModel for a single tab's file list, sort state, and navigation scope.
/// Each tab has its own independent instance with its own FileSystemWatcher and navigation history.
/// </summary>
public sealed partial class TabContentViewModel : ObservableObject, IDisposable
{
    private readonly IFileSystemService _fileSystemService;
    private readonly IShellIntegrationService _shellService;
    private readonly ICloudStatusService _cloudStatusService;
    private readonly IPreviewService _previewService;
    private readonly ILogger<TabContentViewModel> _logger;
    private IDisposable? _watcherHandle;
    private CancellationTokenSource? _loadCts;

    /// <summary>The underlying tab state.</summary>
    public TabState TabState { get; }

    /// <summary>Items displayed in the file list.</summary>
    public ObservableCollection<FileItemViewModel> Items { get; } = [];

    /// <summary>Currently selected items.</summary>
    public ObservableCollection<FileItemViewModel> SelectedItems { get; } = [];

    /// <summary>Navigation history for this tab.</summary>
    public NavigationHistory History => TabState.History;

    [ObservableProperty]
    private string _currentPath = string.Empty;

    [ObservableProperty]
    private string _tabTitle = "New tab";

    [ObservableProperty]
    private bool _isPinned;

    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private string _sortColumn = "Name";

    [ObservableProperty]
    private bool _sortAscending = true;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    private string _freeSpaceText = string.Empty;

    [ObservableProperty]
    private bool _isEmpty;

    [ObservableProperty]
    private FileItemViewModel? _previewItem;

    [ObservableProperty]
    private PreviewContent? _previewContent;

    [ObservableProperty]
    private bool _isAutoRefreshPaused;

    /// <summary>Raised when navigation occurs (for tree sync).</summary>
    public event EventHandler<string>? Navigated;

    /// <summary>Initializes a new instance of the <see cref="TabContentViewModel"/> class.</summary>
    public TabContentViewModel(
        IFileSystemService fileSystemService,
        IShellIntegrationService shellService,
        ICloudStatusService cloudStatusService,
        IPreviewService previewService,
        ILogger<TabContentViewModel> logger,
        TabState? existingState = null)
    {
        _fileSystemService = fileSystemService;
        _shellService = shellService;
        _cloudStatusService = cloudStatusService;
        _previewService = previewService;
        _logger = logger;
        TabState = existingState ?? new TabState();
    }

    /// <summary>Navigates to the specified directory path.</summary>
    [RelayCommand]
    public Task NavigateAsync(string path) => NavigateCoreAsync(path, skipHistory: false);

    /// <summary>Navigates to the specified directory path, optionally skipping history push.</summary>
    private async Task NavigateCoreAsync(string path, bool skipHistory)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        // Resolve shell paths
        if (path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            var resolved = _shellService.ResolveShellPath(path);
            if (resolved is null) return;
            path = resolved;
        }

        // Handle \\?\ prefix for long paths
        path = path.TrimEnd('\\');
        if (!Directory.Exists(path) && !path.Equals("shell:ThisPC", StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning("Path does not exist: {Path}", path);
            return;
        }

        _loadCts?.Cancel();
        _loadCts = new CancellationTokenSource();
        var ct = _loadCts.Token;

        try
        {
            IsLoading = true;
            Items.Clear();

            // Update state
            CurrentPath = path;
            TabState.CurrentPath = path;
            TabTitle = Path.GetFileName(path);
            if (string.IsNullOrEmpty(TabTitle)) TabTitle = path; // Root drive
            TabState.Title = TabTitle;

            // Add to navigation history
            if (!skipHistory)
                History.Navigate(new DirectoryEntry { Path = path, DisplayName = TabTitle });

            // Stop previous watcher
            _watcherHandle?.Dispose();

            // Start enumeration
            var items = new List<FileItemViewModel>();
            await foreach (var entry in _fileSystemService.EnumerateDirectoryAsync(path, ct))
            {
                ct.ThrowIfCancellationRequested();
                entry.TypeDescription = _shellService.GetTypeDescription(entry.FullPath);
                entry.IconIndex = _shellService.GetIconIndex(entry.FullPath, entry.IsDirectory);
                items.Add(new FileItemViewModel(entry));
            }

            // Sort and populate
            var sorted = ApplySort(items);
            foreach (var item in sorted)
            {
                Items.Add(item);
            }

            IsEmpty = Items.Count == 0;
            UpdateStatusText();
            UpdateFreeSpace();

            // Start watcher
            StartWatcher(path);

            // Batch-resolve cloud statuses asynchronously
            if (_cloudStatusService.IsSyncRootDetected)
            {
                _ = ResolveCloudStatusesAsync(ct);
            }

            Navigated?.Invoke(this, path);
        }
        catch (OperationCanceledException)
        {
            // Navigation was superseded
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to navigate to {Path}", path);
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>Navigates back in history.</summary>
    [RelayCommand]
    public async Task GoBackAsync()
    {
        var entry = History.GoBack();
        if (entry is not null) await NavigateCoreAsync(entry.Path, skipHistory: true);
    }

    /// <summary>Navigates forward in history.</summary>
    [RelayCommand]
    public async Task GoForwardAsync()
    {
        var entry = History.GoForward();
        if (entry is not null) await NavigateCoreAsync(entry.Path, skipHistory: true);
    }

    /// <summary>Navigates to the parent directory.</summary>
    [RelayCommand]
    public async Task GoUpAsync()
    {
        var parent = Path.GetDirectoryName(CurrentPath);
        if (parent is not null) await NavigateAsync(parent);
    }

    /// <summary>Refreshes the current directory.</summary>
    [RelayCommand]
    public async Task RefreshAsync()
    {
        if (!string.IsNullOrEmpty(CurrentPath))
            await NavigateAsync(CurrentPath);
    }

    /// <summary>Opens the selected item (navigate into folder or launch file).</summary>
    [RelayCommand]
    public async Task OpenItemAsync(FileItemViewModel? item)
    {
        if (item is null && SelectedItems.Count > 0)
            item = SelectedItems[0];
        if (item is null) return;

        if (item.Entry.IsDirectory)
        {
            await NavigateAsync(item.Entry.FullPath);
        }
        else
        {
            _shellService.OpenFile(item.Entry.FullPath);
        }
    }

    /// <summary>Starts inline rename mode on the selected item.</summary>
    [RelayCommand]
    public void StartRename()
    {
        if (SelectedItems.Count == 1)
        {
            SelectedItems[0].IsRenaming = true;
        }
    }

    /// <summary>Commits a rename operation with optimistic UI.</summary>
    [RelayCommand]
    public async Task CommitRenameAsync(FileItemViewModel item)
    {
        if (!item.IsRenaming) return;

        var oldName = item.Entry.Name;
        var newName = item.EditingName;
        item.IsRenaming = false;

        if (string.IsNullOrWhiteSpace(newName) || newName == oldName) return;

        // Optimistic: update name immediately
        item.Entry.Name = newName;
        item.NotifyNameChanged();

        var success = await _fileSystemService.RenameAsync(item.Entry.FullPath, newName);
        if (!success)
        {
            // Rollback
            item.Entry.Name = oldName;
            item.NotifyNameChanged();
            item.HasError = true;
            item.ErrorMessage = "Rename failed";
        }
    }

    /// <summary>Creates a new folder in the current directory.</summary>
    [RelayCommand]
    public async Task CreateNewFolderAsync()
    {
        if (string.IsNullOrEmpty(CurrentPath)) return;

        try
        {
            var newPath = await _fileSystemService.CreateNewFolderAsync(CurrentPath);
            var name = Path.GetFileName(newPath);

            // Optimistic: add to list immediately in rename mode
            var entry = new FileSystemEntry
            {
                FullPath = newPath,
                Name = name,
                IsDirectory = true,
                DateModified = DateTimeOffset.Now,
                Attributes = FileAttributes.Directory
            };

            var vm = new FileItemViewModel(entry) { IsRenaming = true, EditingName = name, IsNewItem = true };
            Items.Insert(0, vm);
            SelectedItems.Clear();
            SelectedItems.Add(vm);
            UpdateStatusText();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to create new folder");
        }
    }

    /// <summary>Deletes selected items to the Recycle Bin.</summary>
    [RelayCommand]
    public async Task DeleteSelectedAsync()
    {
        var items = SelectedItems.ToList();
        if (items.Count == 0) return;

        // Optimistic: remove immediately
        foreach (var item in items)
        {
            Items.Remove(item);
        }

        var paths = items.Select(i => i.Entry.FullPath).ToList();
        var success = await _fileSystemService.DeleteToRecycleBinAsync(paths);

        if (!success)
        {
            // Rollback
            foreach (var item in items)
            {
                Items.Add(item);
            }
            _logger.LogWarning("Delete to Recycle Bin failed, items restored");
        }

        UpdateStatusText();
    }

    /// <summary>Permanently deletes selected items.</summary>
    [RelayCommand]
    public async Task PermanentDeleteSelectedAsync()
    {
        var items = SelectedItems.ToList();
        if (items.Count == 0) return;

        foreach (var item in items) Items.Remove(item);

        var paths = items.Select(i => i.Entry.FullPath).ToList();
        var success = await _fileSystemService.PermanentDeleteAsync(paths);

        if (!success)
        {
            foreach (var item in items) Items.Add(item);
        }

        UpdateStatusText();
    }

    /// <summary>Copies selected item paths to clipboard.</summary>
    [RelayCommand]
    public void CopyPathToClipboard()
    {
        // The view will read from SelectedItems and set the clipboard
    }

    /// <summary>Selects all items.</summary>
    [RelayCommand]
    public void SelectAll()
    {
        SelectedItems.Clear();
        foreach (var item in Items)
        {
            SelectedItems.Add(item);
        }
        UpdateStatusText();
    }

    /// <summary>Updates the sort column and direction.</summary>
    [RelayCommand]
    public void SetSort(string column)
    {
        if (SortColumn == column)
        {
            SortAscending = !SortAscending;
        }
        else
        {
            SortColumn = column;
            SortAscending = true;
        }

        TabState.SortColumn = SortColumn;
        TabState.SortAscending = SortAscending;

        var sorted = ApplySort(Items.ToList());
        Items.Clear();
        foreach (var item in sorted) Items.Add(item);
    }

    /// <summary>Loads preview for the specified item.</summary>
    public async Task LoadPreviewAsync(FileItemViewModel? item)
    {
        if (item is null || item.Entry.IsDirectory)
        {
            PreviewContent = null;
            PreviewItem = null;
            return;
        }

        PreviewItem = item;
        PreviewContent = await _previewService.GetPreviewAsync(item.Entry.FullPath);
    }

    #region Private Helpers

    private List<FileItemViewModel> ApplySort(List<FileItemViewModel> items)
    {
        // Folders always first
        IOrderedEnumerable<FileItemViewModel> ordered = SortColumn switch
        {
            "Name" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase)
                : items.OrderByDescending(i => i.Entry.IsDirectory).ThenByDescending(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase),
            "DateModified" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.DateModified)
                : items.OrderByDescending(i => i.Entry.IsDirectory).ThenByDescending(i => i.Entry.DateModified),
            "Type" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.TypeDescription, StringComparer.OrdinalIgnoreCase)
                : items.OrderByDescending(i => i.Entry.IsDirectory).ThenByDescending(i => i.Entry.TypeDescription, StringComparer.OrdinalIgnoreCase),
            "Size" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.Size)
                : items.OrderByDescending(i => i.Entry.IsDirectory).ThenByDescending(i => i.Entry.Size),
            _ => items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase)
        };

        return ordered.ToList();
    }

    private void StartWatcher(string path)
    {
        try
        {
            _watcherHandle = _fileSystemService.WatchDirectory(path, change =>
            {
                // Marshal to UI thread — the view will need to call DispatcherQueue.TryEnqueue
                HandleFileSystemChange(change);
            });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to start FileSystemWatcher for {Path}", path);
            IsAutoRefreshPaused = true;
        }
    }

    private void HandleFileSystemChange(FileSystemChangeEvent change)
    {
        switch (change.ChangeType)
        {
            case WatcherChangeTypes.Created:
                var existing = Items.FirstOrDefault(i => i.Entry.FullPath.Equals(change.FullPath, StringComparison.OrdinalIgnoreCase));
                if (existing is null)
                {
                    try
                    {
                        var info = new FileInfo(change.FullPath);
                        var isDir = Directory.Exists(change.FullPath);
                        var entry = new FileSystemEntry
                        {
                            FullPath = change.FullPath,
                            Name = change.Name,
                            IsDirectory = isDir,
                            Size = isDir ? 0 : (info.Exists ? info.Length : 0),
                            DateModified = info.Exists ? info.LastWriteTime : DateTimeOffset.Now,
                            Attributes = info.Exists ? info.Attributes : FileAttributes.Normal
                        };
                        entry.TypeDescription = _shellService.GetTypeDescription(change.FullPath);
                        Items.Add(new FileItemViewModel(entry) { IsNewItem = true });
                        UpdateStatusText();
                    }
                    catch { /* Race condition — file may already be deleted */ }
                }
                break;

            case WatcherChangeTypes.Deleted:
                var toRemove = Items.FirstOrDefault(i => i.Entry.FullPath.Equals(change.FullPath, StringComparison.OrdinalIgnoreCase));
                if (toRemove is not null)
                {
                    Items.Remove(toRemove);
                    UpdateStatusText();
                }
                break;

            case WatcherChangeTypes.Renamed:
                var renamed = Items.FirstOrDefault(i =>
                    change.OldFullPath is not null &&
                    i.Entry.FullPath.Equals(change.OldFullPath, StringComparison.OrdinalIgnoreCase));
                if (renamed is not null)
                {
                    renamed.Entry.Name = change.Name;
                    renamed.NotifyNameChanged();
                }
                break;
        }
    }

    private async Task ResolveCloudStatusesAsync(CancellationToken ct)
    {
        try
        {
            var paths = Items.Select(i => i.Entry.FullPath).ToList();
            var statuses = await _cloudStatusService.GetCloudStatusBatchAsync(paths, ct);
            foreach (var item in Items)
            {
                if (statuses.TryGetValue(item.Entry.FullPath, out var status))
                {
                    item.CloudStatus = status;
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to resolve cloud statuses");
        }
    }

    private void UpdateStatusText()
    {
        if (SelectedItems.Count > 0)
        {
            var totalSize = SelectedItems.Where(i => !i.Entry.IsDirectory).Sum(i => i.Entry.Size);
            StatusText = $"{SelectedItems.Count} item{(SelectedItems.Count != 1 ? "s" : "")} selected, {FormatSize(totalSize)}";
        }
        else
        {
            StatusText = $"{Items.Count} item{(Items.Count != 1 ? "s" : "")}";
        }
    }

    private void UpdateFreeSpace()
    {
        try
        {
            var root = Path.GetPathRoot(CurrentPath);
            if (root is not null)
            {
                var drive = new DriveInfo(root);
                if (drive.IsReady)
                {
                    var freeGb = drive.AvailableFreeSpace / (1024.0 * 1024 * 1024);
                    var totalGb = drive.TotalSize / (1024.0 * 1024 * 1024);
                    FreeSpaceText = $"Free space: {freeGb:F1} GB of {totalGb:F1} GB";
                }
            }
        }
        catch
        {
            FreeSpaceText = string.Empty;
        }
    }

    internal static string FormatSize(long bytes)
    {
        return bytes switch
        {
            < 1024 => $"{bytes} B",
            < 1024 * 1024 => $"{bytes / 1024.0:F1} KB",
            < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):F1} MB",
            _ => $"{bytes / (1024.0 * 1024 * 1024):F2} GB"
        };
    }

    #endregion

    /// <inheritdoc/>
    public void Dispose()
    {
        _loadCts?.Cancel();
        _loadCts?.Dispose();
        _watcherHandle?.Dispose();
    }
}

/// <summary>
/// ViewModel for a single file or folder item in the details list.
/// </summary>
public sealed partial class FileItemViewModel : ObservableObject
{
    /// <summary>The underlying file system entry.</summary>
    public FileSystemEntry Entry { get; }

    /// <summary>Display name.</summary>
    public string Name => Entry.Name;

    /// <summary>Formatted file size.</summary>
    public string SizeText => Entry.IsDirectory ? string.Empty : TabContentViewModel.FormatSize(Entry.Size);

    /// <summary>Formatted modification date.</summary>
    public string DateModifiedText => Entry.DateModified.LocalDateTime.ToString("g");

    /// <summary>Shell type description.</summary>
    public string TypeDescription => Entry.TypeDescription;

    /// <summary>Full path for tooltips.</summary>
    public string FullPath => Entry.FullPath;

    /// <summary>Whether this is a directory.</summary>
    public bool IsDirectory => Entry.IsDirectory;

    /// <summary>Icon glyph based on file type.</summary>
    public string IconGlyph => Entry.IsDirectory ? "\uE8B7" : GetFileIconGlyph(Entry.Extension);

    [ObservableProperty]
    private CloudFileStatus _cloudStatus = CloudFileStatus.NotApplicable;

    [ObservableProperty]
    private bool _isRenaming;

    [ObservableProperty]
    private string _editingName = string.Empty;

    [ObservableProperty]
    private bool _hasError;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private bool _isNewItem;

    [ObservableProperty]
    private bool _isSelected;

    /// <summary>Cloud status icon based on OneDrive state.</summary>
    public string CloudStatusIcon => CloudStatus switch
    {
        CloudFileStatus.CloudOnly => "\uE753",      // Cloud icon
        CloudFileStatus.LocallyAvailable => "\uE73E", // Checkmark
        CloudFileStatus.Syncing => "\uE895",         // Sync icon
        CloudFileStatus.SyncError => "\uE7BA",       // Warning
        CloudFileStatus.SharedWithMe => "\uE902",    // People icon
        CloudFileStatus.Pending => "\uE712",         // Progress ring
        _ => string.Empty
    };

    /// <summary>Initializes a new instance of the <see cref="FileItemViewModel"/> class.</summary>
    public FileItemViewModel(FileSystemEntry entry)
    {
        Entry = entry;
        EditingName = entry.Name;
        CloudStatus = entry.CloudStatus;
    }

    /// <summary>Notifies that the Name property has changed (called externally after Entry.Name mutation).</summary>
    public void NotifyNameChanged()
    {
        OnPropertyChanged(nameof(Name));
    }

    partial void OnCloudStatusChanged(CloudFileStatus value)
    {
        OnPropertyChanged(nameof(CloudStatusIcon));
    }

    private static string GetFileIconGlyph(string extension)
    {
        return extension.ToLowerInvariant() switch
        {
            ".txt" or ".md" or ".log" => "\uE8A5",
            ".pdf" => "\uEA90",
            ".doc" or ".docx" => "\uE8A5",
            ".xls" or ".xlsx" => "\uE80A",
            ".ppt" or ".pptx" => "\uE8A5",
            ".zip" or ".rar" or ".7z" => "\uE8D4",
            ".jpg" or ".jpeg" or ".png" or ".gif" or ".bmp" => "\uEB9F",
            ".mp3" or ".wav" or ".flac" => "\uEC4F",
            ".mp4" or ".avi" or ".mkv" => "\uE714",
            ".exe" or ".msi" => "\uE756",
            ".dll" => "\uE74C",
            ".cs" or ".py" or ".js" or ".ts" => "\uE943",
            _ => "\uE7C3"
        };
    }
}
