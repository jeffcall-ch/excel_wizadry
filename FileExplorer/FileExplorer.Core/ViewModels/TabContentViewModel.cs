// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
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
    private const int IdentityApplyBatchSize = 50;
    private static readonly TimeSpan IdentityApplyDebounce = TimeSpan.FromMilliseconds(250);

    private static readonly Regex RevisionFromFileNameRegex = new(
        @"(?<!\d)(?<rev>\d+\.\d+)(?!\d)",
        RegexOptions.Compiled);

    private readonly IFileSystemService _fileSystemService;
    private readonly IShellIntegrationService _shellService;
    private readonly ICloudStatusService _cloudStatusService;
    private readonly IPreviewService _previewService;
    private readonly IIdentityExtractionService _identityExtractionService;
    private readonly IIdentityCacheService _identityCacheService;
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

    [ObservableProperty]
    private bool _isIdentityHydrationActive;

    /// <summary>Raised when navigation occurs (for tree sync).</summary>
    public event EventHandler<string>? Navigated;

    /// <summary>Raised when the UI should reveal and focus a specific item.</summary>
    public event EventHandler<FileItemViewModel>? RevealItemRequested;

    /// <summary>Initializes a new instance of the <see cref="TabContentViewModel"/> class.</summary>
    public TabContentViewModel(
        IFileSystemService fileSystemService,
        IShellIntegrationService shellService,
        ICloudStatusService cloudStatusService,
        IPreviewService previewService,
        IIdentityExtractionService identityExtractionService,
        IIdentityCacheService identityCacheService,
        ILogger<TabContentViewModel> logger,
        TabState? existingState = null)
    {
        _fileSystemService = fileSystemService;
        _shellService = shellService;
        _cloudStatusService = cloudStatusService;
        _previewService = previewService;
        _identityExtractionService = identityExtractionService;
        _identityCacheService = identityCacheService;
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

        try
        {
            path = NormalizeNavigationPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            _logger.LogWarning(ex, "Path is invalid: {Path}", path);
            return;
        }

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
            TabTitle = BuildFriendlyTabTitle(path);
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

            _ = HydrateIdentityColumnsAsync(ct);

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

        if (string.Equals(item.Entry.Extension, ".lnk", StringComparison.OrdinalIgnoreCase))
        {
            var shortcutTarget = _shellService.ResolveShortcutTarget(item.Entry.FullPath);
            if (!string.IsNullOrWhiteSpace(shortcutTarget))
            {
                if (Directory.Exists(shortcutTarget))
                {
                    await NavigateAsync(shortcutTarget);
                    return;
                }

                if (File.Exists(shortcutTarget))
                {
                    _shellService.OpenFile(shortcutTarget);
                    return;
                }
            }
        }

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
            SelectedItems[0].EditingName = SelectedItems[0].Name;
            RevealItemRequested?.Invoke(this, SelectedItems[0]);
        }
    }

    /// <summary>Commits a rename operation with optimistic UI.</summary>
    [RelayCommand]
    public async Task CommitRenameAsync(FileItemViewModel item)
    {
        if (!item.IsRenaming) return;

        var oldName = item.Entry.Name;
        var oldFullPath = item.Entry.FullPath;
        var newName = item.EditingName;
        item.IsRenaming = false;
        item.IsNewItem = false;

        if (string.IsNullOrWhiteSpace(newName) || newName == oldName)
        {
            item.EditingName = oldName;
            MoveItemToSortedPosition(item);
            RevealItemRequested?.Invoke(this, item);
            return;
        }

        // Optimistic: update name immediately
        item.Entry.Name = newName;
        item.NotifyNameChanged();

        var success = await _fileSystemService.RenameAsync(oldFullPath, newName);
        if (!success)
        {
            // Rollback
            item.Entry.Name = oldName;
            item.EditingName = oldName;
            item.NotifyNameChanged();
            item.HasError = true;
            item.ErrorMessage = "Rename failed";
            RevealItemRequested?.Invoke(this, item);
            return;
        }

        var parentPath = Path.GetDirectoryName(oldFullPath);
        var renamedPath = string.IsNullOrEmpty(parentPath)
            ? newName
            : Path.Combine(parentPath, newName);

        item.Entry.FullPath = renamedPath;
        item.Entry.Extension = Path.GetExtension(newName);
        item.Entry.TypeDescription = _shellService.GetTypeDescription(renamedPath);
        item.Entry.IconIndex = _shellService.GetIconIndex(renamedPath, item.Entry.IsDirectory);
        item.NotifyRenamed();

        MoveItemToSortedPosition(item);
        RevealItemRequested?.Invoke(this, item);
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
            RevealItemRequested?.Invoke(this, vm);
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
                ? items.OrderBy(i => i.Entry.DateModified).ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase)
                : items.OrderByDescending(i => i.Entry.DateModified).ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase),
            "Type" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.TypeDescription, StringComparer.OrdinalIgnoreCase)
                : items.OrderByDescending(i => i.Entry.IsDirectory).ThenByDescending(i => i.Entry.TypeDescription, StringComparer.OrdinalIgnoreCase),
            "ExtractedProject" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory)
                    .ThenBy(i => i.ExtractedProject, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase)
                : items.OrderByDescending(i => i.Entry.IsDirectory)
                    .ThenByDescending(i => i.ExtractedProject, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase),
            "ExtractedRevision" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory)
                    .ThenBy(i => i.ExtractedRevision, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase)
                : items.OrderByDescending(i => i.Entry.IsDirectory)
                    .ThenByDescending(i => i.ExtractedRevision, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase),
            "Size" => SortAscending
                ? items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.Size)
                : items.OrderByDescending(i => i.Entry.IsDirectory).ThenByDescending(i => i.Entry.Size),
            _ => items.OrderByDescending(i => i.Entry.IsDirectory).ThenBy(i => i.Entry.Name, StringComparer.OrdinalIgnoreCase)
        };

        return ordered.ToList();
    }

    private void MoveItemToSortedPosition(FileItemViewModel item)
    {
        var currentIndex = Items.IndexOf(item);
        if (currentIndex < 0) return;

        var sortedItems = ApplySort(Items.ToList());
        var sortedIndex = sortedItems.IndexOf(item);
        if (sortedIndex < 0 || sortedIndex == currentIndex) return;

        Items.Move(currentIndex, sortedIndex);
    }

    public static string NormalizeNavigationPath(string path)
    {
        path = StripLongPathPrefix(path.Trim().Trim('"'));

        if (path.Length == 2 && char.IsLetter(path[0]) && path[1] == Path.VolumeSeparatorChar)
            return path + Path.DirectorySeparatorChar;

        var fullPath = Path.GetFullPath(path);
        fullPath = StripLongPathPrefix(fullPath);
        var root = Path.GetPathRoot(fullPath);

        if (!string.IsNullOrEmpty(root) && fullPath.Equals(root, StringComparison.OrdinalIgnoreCase))
            return root;

        return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private static string StripLongPathPrefix(string path)
    {
        if (string.IsNullOrEmpty(path))
            return path;

        if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            return @"\\" + path[8..];

        if (path.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            return path[4..];

        return path;
    }

    private static string BuildFriendlyTabTitle(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return "New tab";

        var normalized = NormalizeNavigationPath(path);
        var root = Path.GetPathRoot(normalized);

        if (!string.IsNullOrEmpty(root) && string.Equals(normalized, root, StringComparison.OrdinalIgnoreCase))
        {
            return root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        var leaf = Path.GetFileName(normalized.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        if (!string.IsNullOrWhiteSpace(leaf))
            return leaf;

        return normalized;
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
                        var attributes = File.Exists(change.FullPath) || isDir
                            ? File.GetAttributes(change.FullPath)
                            : FileAttributes.Normal;

                        var entry = new FileSystemEntry
                        {
                            FullPath = change.FullPath,
                            Name = change.Name,
                            IsDirectory = isDir,
                            Size = isDir ? 0 : (info.Exists ? info.Length : 0),
                            DateModified = info.Exists ? info.LastWriteTime : DateTimeOffset.Now,
                            Attributes = attributes,
                            Extension = isDir ? string.Empty : Path.GetExtension(change.FullPath)
                        };
                        entry.TypeDescription = _shellService.GetTypeDescription(change.FullPath);
                        var vm = new FileItemViewModel(entry) { IsNewItem = true };
                        Items.Add(vm);

                        if (_cloudStatusService.IsSyncRootDetected)
                        {
                            _ = RefreshSingleCloudStatusAsync(vm);
                        }

                        if (!isDir &&
                            _identityExtractionService.SupportsExtension(entry.Extension) &&
                            !ShouldSkipIndexing(entry))
                        {
                            _ = RefreshIdentityForSingleFileAsync(vm);
                        }

                        UpdateStatusText();
                    }
                    catch { /* Race condition — file may already be deleted */ }
                }
                break;

            case WatcherChangeTypes.Changed:
                var changed = Items.FirstOrDefault(i => i.Entry.FullPath.Equals(change.FullPath, StringComparison.OrdinalIgnoreCase));
                if (changed is not null && _cloudStatusService.IsSyncRootDetected)
                {
                    _ = RefreshSingleCloudStatusAsync(changed);
                }

                if (changed is not null &&
                    !changed.Entry.IsDirectory &&
                    _identityExtractionService.SupportsExtension(changed.Entry.Extension) &&
                    !ShouldSkipIndexing(changed.Entry))
                {
                    _ = RefreshIdentityForSingleFileAsync(changed);
                }
                break;

            case WatcherChangeTypes.Deleted:
                var toRemove = Items.FirstOrDefault(i => i.Entry.FullPath.Equals(change.FullPath, StringComparison.OrdinalIgnoreCase));
                if (toRemove is not null)
                {
                    Items.Remove(toRemove);
                    UpdateStatusText();
                }

                _ = _identityCacheService.DeleteByNormalizedPathAsync(change.FullPath);
                break;

            case WatcherChangeTypes.Renamed:
                var renamed = Items.FirstOrDefault(i =>
                    change.OldFullPath is not null &&
                    i.Entry.FullPath.Equals(change.OldFullPath, StringComparison.OrdinalIgnoreCase));
                if (renamed is not null)
                {
                    var previousPath = renamed.Entry.FullPath;
                    renamed.Entry.Name = change.Name;
                    renamed.Entry.FullPath = change.FullPath;
                    renamed.Entry.Extension = Path.GetExtension(change.FullPath);
                    renamed.Entry.TypeDescription = _shellService.GetTypeDescription(change.FullPath);
                    renamed.NotifyNameChanged();
                    renamed.NotifyRenamed();

                    _ = _identityCacheService.DeleteByNormalizedPathAsync(previousPath);
                    _ = _identityCacheService.DeleteByNormalizedPathAsync(change.FullPath);

                    if (!renamed.Entry.IsDirectory &&
                        _identityExtractionService.SupportsExtension(renamed.Entry.Extension) &&
                        !ShouldSkipIndexing(renamed.Entry))
                    {
                        _ = RefreshIdentityForSingleFileAsync(renamed);
                    }
                }
                break;
        }
    }

    private async Task ResolveCloudStatusesAsync(CancellationToken ct)
    {
        try
        {
            var entries = Items.Select(i => i.Entry).ToList();
            var statuses = await _cloudStatusService.GetCloudStatusBatchAsync(entries, ct);
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

    private async Task HydrateIdentityColumnsAsync(CancellationToken ct)
    {
        try
        {
            IsIdentityHydrationActive = true;

            var candidates = Items
                .Where(item =>
                    !item.Entry.IsDirectory &&
                    _identityExtractionService.SupportsExtension(item.Entry.Extension) &&
                    !ShouldSkipIndexing(item.Entry))
                .ToList();

            if (candidates.Count == 0)
                return;

            var parentPath = _identityCacheService.NormalizePath(CurrentPath);
            var cacheEntries = await _identityCacheService
                .GetEntriesByParentPathAsync(parentPath, ct);

            var liveEntries = new List<(FileItemViewModel Item, string FileKey, long Size, long LastWriteTicksUtc)>(candidates.Count);
            var liveKeySet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var candidate in candidates)
            {
                var fileKey = _identityCacheService.BuildFileKey(
                    candidate.Entry.FullPath,
                    candidate.Entry.Size,
                    candidate.Entry.DateModified);

                liveEntries.Add((
                    candidate,
                    fileKey,
                    candidate.Entry.Size,
                    candidate.Entry.DateModified.UtcTicks));

                liveKeySet.Add(fileKey);
            }

            var orphanKeys = cacheEntries.Keys
                .Where(key => !liveKeySet.Contains(key))
                .ToList();

            if (orphanKeys.Count > 0)
            {
                await _identityCacheService.DeleteByFileKeysAsync(orphanKeys, ct);
            }

            var cacheApplyQueue = new List<(FileItemViewModel Item, DocumentIdentity Identity)>(liveEntries.Count);
            var extractionQueue = new List<(FileItemViewModel Item, string FileKey, long Size, long LastWriteTicksUtc)>(liveEntries.Count);

            foreach (var liveEntry in liveEntries)
            {
                var revisionFromFileName = ExtractRevisionFromFileName(liveEntry.Item.Entry.Name);
                var isPdf = string.Equals(liveEntry.Item.Entry.Extension, ".pdf", StringComparison.OrdinalIgnoreCase);

                if (cacheEntries.TryGetValue(liveEntry.FileKey, out var cached))
                {
                    var identity = cached.ToDocumentIdentity();

                    if (!string.IsNullOrWhiteSpace(revisionFromFileName) &&
                        (!isPdf || string.IsNullOrWhiteSpace(identity.Revision)))
                    {
                        identity = identity with { Revision = revisionFromFileName };
                    }

                    if (identity.HasAnyValue && NeedsIdentityUpdate(liveEntry.Item, identity))
                    {
                        cacheApplyQueue.Add((liveEntry.Item, identity));
                    }

                    continue;
                }

                extractionQueue.Add(liveEntry);
            }

            if (cacheApplyQueue.Count > 0)
            {
                await ApplyIdentityUpdatesBufferedAsync(cacheApplyQueue, ct);
            }

            if (extractionQueue.Count == 0)
            {
                return;
            }

            var filePaths = extractionQueue
                .Select(entry => entry.Item.Entry.FullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var stopwatch = Stopwatch.StartNew();
            var identities = await _identityExtractionService.ExtractIdentityBatchAsync(filePaths, ct);
            stopwatch.Stop();

            var applyQueue = new List<(FileItemViewModel Item, DocumentIdentity Identity)>(extractionQueue.Count);
            var upsertQueue = new List<FileIdentityCacheRecord>(extractionQueue.Count);

            foreach (var extractionEntry in extractionQueue)
            {
                var item = extractionEntry.Item;
                var revisionFromFileName = ExtractRevisionFromFileName(item.Entry.Name);
                var isPdf = string.Equals(item.Entry.Extension, ".pdf", StringComparison.OrdinalIgnoreCase);

                DocumentIdentity identity = DocumentIdentity.Empty;
                if (identities.TryGetValue(item.Entry.FullPath, out var extractedIdentity))
                {
                    identity = extractedIdentity;
                }

                if (!string.IsNullOrWhiteSpace(revisionFromFileName) &&
                    (!isPdf || string.IsNullOrWhiteSpace(identity.Revision)))
                {
                    identity = identity with { Revision = revisionFromFileName };
                }

                if (!identity.HasAnyValue)
                    continue;

                if (NeedsIdentityUpdate(item, identity))
                {
                    applyQueue.Add((item, identity));
                }

                upsertQueue.Add(new FileIdentityCacheRecord
                {
                    FileKey = extractionEntry.FileKey,
                    ParentPath = parentPath,
                    ProjectTitle = identity.ProjectTitle,
                    Revision = identity.Revision,
                    DocumentName = identity.DocumentName,
                    LastWriteTime = extractionEntry.LastWriteTicksUtc,
                    FileSize = extractionEntry.Size
                });
            }

            if (applyQueue.Count > 0)
            {
                await ApplyIdentityUpdatesBufferedAsync(applyQueue, ct);
            }

            if (upsertQueue.Count > 0)
            {
                await _identityCacheService.UpsertEntriesAsync(upsertQueue, ct);
            }

            _logger.LogInformation(
                "Identity hydration elapsed: {ElapsedMs} ms for {Count} candidate files in {Path}",
                stopwatch.ElapsedMilliseconds,
                candidates.Count,
                CurrentPath);
        }
        catch (OperationCanceledException)
        {
            // Navigation changed; ignore stale identity work.
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Identity hydration failed for {Path}", CurrentPath);
        }
        finally
        {
            IsIdentityHydrationActive = false;
        }
    }

    private async Task ApplyIdentityUpdatesBufferedAsync(
        IReadOnlyList<(FileItemViewModel Item, DocumentIdentity Identity)> updates,
        CancellationToken ct)
    {
        for (var index = 0; index < updates.Count; index += IdentityApplyBatchSize)
        {
            ct.ThrowIfCancellationRequested();
            await Task.Delay(IdentityApplyDebounce, ct);

            var count = Math.Min(IdentityApplyBatchSize, updates.Count - index);
            for (var offset = 0; offset < count; offset++)
            {
                var update = updates[index + offset];
                update.Item.ApplyIdentity(update.Identity);
            }
        }
    }

    private async Task RefreshIdentityForSingleFileAsync(FileItemViewModel item)
    {
        if (item.Entry.IsDirectory || ShouldSkipIndexing(item.Entry))
            return;

        try
        {
            var fileInfo = new FileInfo(item.Entry.FullPath);
            if (!fileInfo.Exists)
                return;

            var fileKey = _identityCacheService.BuildFileKey(
                item.Entry.FullPath,
                fileInfo.Length,
                fileInfo.LastWriteTimeUtc);

            var cached = await _identityCacheService.GetEntryByFileKeyAsync(fileKey);
            if (cached is not null)
            {
                var cachedIdentity = cached.ToDocumentIdentity();
                if (cachedIdentity.HasAnyValue && NeedsIdentityUpdate(item, cachedIdentity))
                {
                    item.ApplyIdentity(cachedIdentity);
                }

                return;
            }

            var identity = await _identityExtractionService.ExtractIdentityAsync(item.Entry.FullPath);
            var revisionFromFileName = ExtractRevisionFromFileName(item.Entry.Name);
            var isPdf = string.Equals(item.Entry.Extension, ".pdf", StringComparison.OrdinalIgnoreCase);

            if (!string.IsNullOrWhiteSpace(revisionFromFileName) &&
                (!isPdf || string.IsNullOrWhiteSpace(identity.Revision)))
            {
                identity = identity with { Revision = revisionFromFileName };
            }

            if (!identity.HasAnyValue)
            {
                await _identityCacheService.DeleteByNormalizedPathAsync(item.Entry.FullPath);
                return;
            }

            if (NeedsIdentityUpdate(item, identity))
            {
                item.ApplyIdentity(identity);
            }

            await _identityCacheService.UpsertEntriesAsync(
                [
                    new FileIdentityCacheRecord
                    {
                        FileKey = fileKey,
                        ParentPath = _identityCacheService.NormalizePath(Path.GetDirectoryName(item.Entry.FullPath) ?? string.Empty),
                        ProjectTitle = identity.ProjectTitle,
                        Revision = identity.Revision,
                        DocumentName = identity.DocumentName,
                        LastWriteTime = fileInfo.LastWriteTimeUtc.Ticks,
                        FileSize = fileInfo.Length
                    }
                ]);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Single-file identity refresh failed for {Path}", item.Entry.FullPath);
        }
    }

    private static bool NeedsIdentityUpdate(FileItemViewModel item, DocumentIdentity identity)
    {
        return !string.Equals(item.ExtractedProject, identity.ProjectTitle, StringComparison.Ordinal) ||
            !string.Equals(item.ExtractedRevision, identity.Revision, StringComparison.Ordinal) ||
            !string.Equals(item.ExtractedDocumentName, identity.DocumentName, StringComparison.Ordinal);
    }

    private static bool ShouldSkipIndexing(FileSystemEntry entry)
    {
        if (entry.IsDirectory)
            return true;

        if (entry.Name.StartsWith("~$", StringComparison.OrdinalIgnoreCase))
            return true;

        if (entry.IsHidden || entry.IsSystem)
            return true;

        return false;
    }

    private static string ExtractRevisionFromFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return string.Empty;

        var nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
        if (string.IsNullOrWhiteSpace(nameWithoutExtension))
            return string.Empty;

        var matches = RevisionFromFileNameRegex.Matches(nameWithoutExtension);
        if (matches.Count == 0)
            return string.Empty;

        var value = matches[^1].Groups["rev"].Value;
        return string.IsNullOrWhiteSpace(value) ? string.Empty : value;
    }

    private async Task RefreshSingleCloudStatusAsync(FileItemViewModel item)
    {
        try
        {
            item.CloudStatus = await _cloudStatusService.GetCloudStatusAsync(item.Entry.FullPath);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to refresh cloud status for {Path}", item.Entry.FullPath);
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

    /// <summary>Returns the current tab title as a fallback header text.</summary>
    public override string ToString() => TabTitle;
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

    /// <summary>File extension including leading dot.</summary>
    public string Extension => Entry.Extension;

    /// <summary>Extracted project title from document anchors.</summary>
    public string ExtractedProject => Entry.ExtractedProject;

    /// <summary>Extracted revision from document anchors.</summary>
    public string ExtractedRevision => Entry.ExtractedRevision;

    /// <summary>Extracted document name from document anchors.</summary>
    public string ExtractedDocumentName => Entry.ExtractedDocumentName;

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
        CloudFileStatus.AlwaysAvailable => "\uE718", // Pinned
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

    /// <summary>Notifies dependent properties that change after a successful rename.</summary>
    public void NotifyRenamed()
    {
        OnPropertyChanged(nameof(Name));
        OnPropertyChanged(nameof(FullPath));
        OnPropertyChanged(nameof(Extension));
        OnPropertyChanged(nameof(TypeDescription));
        OnPropertyChanged(nameof(IconGlyph));
    }

    /// <summary>Applies extracted identity values and raises binding updates.</summary>
    public void ApplyIdentity(DocumentIdentity identity)
    {
        Entry.ExtractedProject = identity.ProjectTitle;
        Entry.ExtractedRevision = identity.Revision;
        Entry.ExtractedDocumentName = identity.DocumentName;

        OnPropertyChanged(nameof(ExtractedProject));
        OnPropertyChanged(nameof(ExtractedRevision));
        OnPropertyChanged(nameof(ExtractedDocumentName));
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
