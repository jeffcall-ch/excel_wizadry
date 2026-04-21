// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Text;
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
    private static readonly TimeSpan IdentityApplyDebounce = TimeSpan.FromMilliseconds(150);

        @"(?<!\d)(?<rev>\d+\.\d+)(?!\d)",
        RegexOptions.Compiled);

        @"^\d+\.\d+$",
        RegexOptions.Compiled);


    private readonly IFileSystemService _fileSystemService;
    private readonly IShellIntegrationService _shellService;
    private readonly ICloudStatusService _cloudStatusService;
    private readonly IPreviewService _previewService;
    private readonly ITabPathStateService _tabPathStateService;
    private readonly ISettingsService _settingsService;
    private readonly IDirectorySnapshotCache _snapshotCache;
    private readonly ILogger<TabContentViewModel> _logger;
    private readonly SynchronizationContext? _uiContext;
    private IDisposable? _watcherHandle;
    private CancellationTokenSource? _loadCts;
    private int _tabPathSyncVersion;
    private DateTimeOffset _lastLoadCompletedUtc = DateTimeOffset.MinValue;

    /// <summary>UTC timestamp of the last completed directory load or incremental refresh.</summary>
    public DateTimeOffset LastLoadCompletedUtc => _lastLoadCompletedUtc;


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

    /// <summary>Raised when a rename operation fails, with the user-readable error message.</summary>
    public event EventHandler<string>? RenameError;

    /// <summary>Raised when the UI should reveal and focus a specific item.</summary>
    public event EventHandler<FileItemViewModel>? RevealItemRequested;

    /// <summary>Initializes a new instance of the <see cref="TabContentViewModel"/> class.</summary>
    public TabContentViewModel(
        IFileSystemService fileSystemService,
        IShellIntegrationService shellService,
        ICloudStatusService cloudStatusService,
        IPreviewService previewService,
        ITabPathStateService tabPathStateService,
        ISettingsService settingsService,
        IDirectorySnapshotCache snapshotCache,
        ILogger<TabContentViewModel> logger,
        TabState? existingState = null)
    {
        _fileSystemService = fileSystemService;
        _shellService = shellService;
        _cloudStatusService = cloudStatusService;
        _previewService = previewService;
        _tabPathStateService = tabPathStateService;
        _settingsService = settingsService;
        _snapshotCache = snapshotCache;
        _logger = logger;
        _uiContext = SynchronizationContext.Current;
        TabState = existingState ?? new TabState();

        if (!string.IsNullOrWhiteSpace(TabState.CurrentPath))
        {
            CurrentPath = TabState.CurrentPath;
        }
    }

    partial void OnCurrentPathChanged(string value)
    {
        TabState.CurrentPath = value;
        var fallbackTitle = BuildFriendlyTabTitle(value);
        if (!string.Equals(TabTitle, fallbackTitle, StringComparison.Ordinal))
        {
            TabTitle = fallbackTitle;
        }

        TabState.Title = TabTitle;

        var version = Interlocked.Increment(ref _tabPathSyncVersion);
        _ = PersistAndApplyTabPathStateAsync(value, fallbackTitle, version);
    }

    private async Task PersistAndApplyTabPathStateAsync(string currentPath, string fallbackTitle, int version)
    {
        try
        {
            await _tabPathStateService.UpsertCurrentPathAsync(TabState.Id, currentPath, fallbackTitle);
            var persisted = await _tabPathStateService.GetByTabIdAsync(TabState.Id);
            if (persisted is null)
                return;

            if (version != Volatile.Read(ref _tabPathSyncVersion))
                return;

            var titleFromDb = string.IsNullOrWhiteSpace(persisted.DisplayTitle)
                ? fallbackTitle
                : persisted.DisplayTitle;

            void ApplyTitleFromDb()
            {
                if (version != Volatile.Read(ref _tabPathSyncVersion))
                    return;

                if (!string.Equals(TabTitle, titleFromDb, StringComparison.Ordinal))
                {
                    TabTitle = titleFromDb;
                }

                TabState.Title = titleFromDb;
            }

            if (_uiContext is null)
                return;

            if (!ReferenceEquals(SynchronizationContext.Current, _uiContext))
            {
                _uiContext.Post(_ => ApplyTitleFromDb(), null);
                return;
            }

            ApplyTitleFromDb();
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to sync tab path state for tab {TabId}", TabState.Id);
        }
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
            // Keep old items visible during the background disk scan (no blank-screen hang).
            // Items.Clear() is deferred until the first batch of new items is ready.

            // Update state
            CurrentPath = path;
            var navSw = System.Diagnostics.Stopwatch.StartNew();

            // Add to navigation history
            if (!skipHistory)
                History.Navigate(new DirectoryEntry { Path = path, DisplayName = TabTitle });

            // Stop previous watcher
            _watcherHandle?.Dispose();

            // Stage 1: stream entries from disk on a background thread.
            // After the first 50 entries arrive the UI is updated immediately (fast first-paint),
            // then the remainder of the scan continues and the final sorted list replaces the
            // partial view in one pass — no waiting for the full scan to show anything.
            System.Diagnostics.Debug.WriteLine($"[Navigate] start scan  path={path}");
            var snapshot = await StreamLoadDirectoryAsync(path, ct);
            System.Diagnostics.Debug.WriteLine($"[Navigate] scan done  {snapshot.Count} items  {navSw.ElapsedMilliseconds} ms  path={path}");
            var items = snapshot.Select(entry => new FileItemViewModel(entry)).ToList();

            ApplyFolderSortPreference(path);

            // Stage 2: apply full sorted list over the partial first-batch already shown.
            var sorted = ApplySort(items);
            ReconcileItemsWithSorted(sorted);
            _lastLoadCompletedUtc = DateTimeOffset.UtcNow;
            System.Diagnostics.Debug.WriteLine($"[Navigate] items shown  {navSw.ElapsedMilliseconds} ms  path={path}");


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

    private async Task<List<FileSystemEntry>> LoadDirectorySnapshotAsync(string path, CancellationToken ct)
    {
        return await Task.Run(async () =>
        {
            var snapshot = new List<FileSystemEntry>(capacity: 256);

            await foreach (var entry in _fileSystemService.EnumerateDirectoryAsync(path, ct))
            {
                ct.ThrowIfCancellationRequested();
                entry.TypeDescription = GetTypeDescriptionFast(entry);
                entry.IconIndex = 0;
                snapshot.Add(entry);

                if (snapshot.Count % 512 == 0)
                    await Task.Yield();
            }

            return snapshot;
        }, ct);
    }

    private async Task AppendItemsInChunksAsync(IReadOnlyList<FileItemViewModel> items, CancellationToken ct, int batchSize)
    {
        if (items.Count == 0)
            return;

        var firstBatchCount = Math.Min(batchSize, items.Count);
        for (var index = 0; index < firstBatchCount; index++)
        {
            ct.ThrowIfCancellationRequested();
            Items.Add(items[index]);
        }

        if (firstBatchCount >= items.Count)
            return;

        for (var index = firstBatchCount; index < items.Count; index++)
        {
            ct.ThrowIfCancellationRequested();
            Items.Add(items[index]);

            if (((index - firstBatchCount) + 1) % batchSize == 0)
                await Task.Yield();
        }
    }

    /// <summary>
    /// Returns directory entries for <paramref name="path"/>.
    /// If a pre-scanned snapshot exists in the cache it is consumed and posted to the UI
    /// instantly (0 ms disk I/O).  Otherwise entries are streamed from disk via Win32
    /// FindFirstFileEx, with the first 50 sent to the UI immediately for fast first-paint.
    /// </summary>
    private async Task<List<FileSystemEntry>> StreamLoadDirectoryAsync(string path, CancellationToken ct)
    {
        // ── Cache hit: snapshot pre-scanned at startup ────────────────────────
        if (_snapshotCache.TryConsume(path, out var cached))
        {
            System.Diagnostics.Debug.WriteLine($"[Navigate] snapshot cache HIT  {cached.Count} items  path={path}");
            var all = cached is List<FileSystemEntry> list ? list : [.. cached];

            // Clear Items synchronously so ReconcileItemsWithSorted (called by the caller
            // immediately after we return, with no yield point) operates on a clean slate.
            // Do NOT use _uiContext.Post here: a posted action would race against
            // ReconcileItemsWithSorted and overwrite the correctly-sorted Items with an
            // unsorted snapshot, breaking date-divider grouping on first load.
            Items.Clear();

            return all;
        }

        // ── Cache miss: stream from disk ──────────────────────────────────────
        System.Diagnostics.Debug.WriteLine($"[Navigate] snapshot cache MISS  path={path}");

        const int firstBatchSize = 50;
        var entries = new List<FileSystemEntry>(capacity: 256);
        var firstBatchPosted = false;

        await Task.Run(async () =>
        {
            await foreach (var entry in _fileSystemService.EnumerateDirectoryAsync(path, ct))
            {
                ct.ThrowIfCancellationRequested();
                entry.TypeDescription = GetTypeDescriptionFast(entry);
                entry.IconIndex = 0;
                entries.Add(entry);

                // As soon as first 50 entries arrive, send them to the UI thread immediately.
                if (!firstBatchPosted && entries.Count >= firstBatchSize)
                {
                    firstBatchPosted = true;
                    var firstBatch = entries
                        .Take(firstBatchSize)
                        .Select(e => new FileItemViewModel(e))
                        .ToList();
                    _uiContext?.Post(_ =>
                    {
                        if (ct.IsCancellationRequested) return;
                        Items.Clear();
                        foreach (var vm in firstBatch) Items.Add(vm);
                    }, null);
                }

                if (entries.Count % 512 == 0)
                    await Task.Yield();
            }
        }, ct);

        return entries;
    }

    /// <summary>
    /// Applies a fully-sorted target list over the current Items collection using
    /// Add/Remove/Move so no row is destroyed needlessly. If Items already contains
    /// a partial first-batch from StreamLoadDirectoryAsync, those rows are reused.
    /// </summary>
    private void ReconcileItemsWithSorted(IReadOnlyList<FileItemViewModel> sorted)
    {
        // Build a path→vm lookup for what is already in Items (the streaming first-batch).
        var existingByPath = new Dictionary<string, FileItemViewModel>(
            Items.Count, StringComparer.OrdinalIgnoreCase);
        foreach (var item in Items)
            existingByPath.TryAdd(item.Entry.FullPath, item);

        // Remove items no longer in the sorted list (should be none on first load).
        for (var i = Items.Count - 1; i >= 0; i--)
        {
            if (!sorted.Any(s => string.Equals(s.Entry.FullPath, Items[i].Entry.FullPath,
                    StringComparison.OrdinalIgnoreCase)))
                Items.RemoveAt(i);
        }

        // Add missing items and move to correct positions.
        for (var i = 0; i < sorted.Count; i++)
        {
            var target = sorted[i];
            var existingIndex = -1;
            for (var j = i; j < Items.Count; j++)
            {
                if (string.Equals(Items[j].Entry.FullPath, target.Entry.FullPath,
                        StringComparison.OrdinalIgnoreCase))
                { existingIndex = j; break; }
            }

            if (existingIndex < 0)
            {
                // New item (arrived after the first-batch or wasn't in first-batch).
                Items.Insert(i, target);
            }
            else if (existingIndex != i)
            {
                Items.Move(existingIndex, i);
            }
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
            await RefreshIncrementalAsync(CurrentPath);
    }

    /// <summary>
    /// Adds new entries to <see cref="Items"/> at their correct sorted positions without
    /// touching the disk.  Used for optimistic destination inserts after a move/copy so
    /// the user sees pasted items instantly instead of waiting for re-enumeration.
    /// </summary>
    public void InsertItemsSorted(IEnumerable<FileSystemEntry> entries)
    {
        foreach (var entry in entries)
        {
            if (string.IsNullOrEmpty(entry.TypeDescription))
                entry.TypeDescription = GetTypeDescriptionFast(entry);

            var vm = new FileItemViewModel(entry);
            Items.Add(vm);
        }

        // Re-position all items according to current sort preference.
        var sorted = ApplySort(Items.ToList());
        ReconcileItemsWithSorted(sorted);

        IsEmpty = Items.Count == 0;
        UpdateStatusText();
    }

    /// <summary>Refreshes the status bar text to reflect the current item / selection count.</summary>
    public void RefreshStatusText() => UpdateStatusText();

    /// <summary>Performs a full reload of the current directory.</summary>
    public async Task ForceReloadAsync()
    {
        if (!string.IsNullOrEmpty(CurrentPath))
            await NavigateCoreAsync(CurrentPath, skipHistory: true);
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

        // Capture the error detail from the OperationError event raised inside RenameAsync.
        string? renameErrorDetail = null;
        void OnOperationError(object? s, FileOperationErrorEventArgs e) => renameErrorDetail = e.ErrorMessage;
        _fileSystemService.OperationError += OnOperationError;
        var success = await _fileSystemService.RenameAsync(oldFullPath, newName);
        _fileSystemService.OperationError -= OnOperationError;

        if (!success)
        {
            // Rollback
            item.Entry.Name = oldName;
            item.EditingName = oldName;
            item.NotifyNameChanged();
            item.HasError = true;
            item.ErrorMessage = renameErrorDetail ?? "Rename failed";
            RenameError?.Invoke(this, item.ErrorMessage);
            RevealItemRequested?.Invoke(this, item);
            return;
        }

        var parentPath = Path.GetDirectoryName(oldFullPath);
        var renamedPath = string.IsNullOrEmpty(parentPath)
            ? newName
            : Path.Combine(parentPath, newName);

        item.Entry.FullPath = renamedPath;
        item.Entry.Extension = Path.GetExtension(newName);
        item.Entry.TypeDescription = GetTypeDescriptionFast(item.Entry);
        item.Entry.IconIndex = 0;
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

        PersistCurrentFolderSortPreference();

        // Apply sort in-place using Move operations so no row is ever destroyed and
        // re-created — eliminates the full-blink caused by Clear + sequential Add.
        ApplySortInPlace();
    }

    private void ApplySortInPlace()
    {
        if (Items.Count <= 1)
            return;

        var target = ApplySort(Items.ToList());
        for (var i = 0; i < target.Count; i++)
        {
            var cur = Items.IndexOf(target[i]);
            if (cur != i)
                Items.Move(cur, i);
        }
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

    private async Task RefreshIncrementalAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || IsLoading)
            return;

        if (Items.Any(item => item.IsRenaming))
            return;

        try
        {
            var refreshSw = System.Diagnostics.Stopwatch.StartNew();
            var latestEntries = await LoadDirectorySnapshotAsync(path, CancellationToken.None);
            System.Diagnostics.Debug.WriteLine($"[Refresh] LoadDirectorySnapshot: {refreshSw.ElapsedMilliseconds} ms  ({latestEntries.Count} entries)  path={path}");

            var comparer = StringComparer.OrdinalIgnoreCase;
            var latestByPath = new Dictionary<string, FileSystemEntry>(comparer);
            foreach (var entry in latestEntries)
            {
                latestByPath.TryAdd(entry.FullPath, entry);
            }

            if (!HasMaterialChanges(latestByPath))
            {
                if (_watcherHandle is null || IsAutoRefreshPaused)
                {
                    _watcherHandle?.Dispose();
                    StartWatcher(path);
                }

                return;
            }

            var existingByPath = new Dictionary<string, FileItemViewModel>(comparer);
            foreach (var item in Items)
            {
                if (!existingByPath.ContainsKey(item.Entry.FullPath))
                {
                    existingByPath[item.Entry.FullPath] = item;
                }
            }

            var selectedPaths = SelectedItems
                .Select(item => item.Entry.FullPath)
                .Where(pathValue => !string.IsNullOrWhiteSpace(pathValue))
                .ToHashSet(comparer);

            var hasChanges = false;
            var affectedItems = new List<FileItemViewModel>();
            var reorderCandidates = new HashSet<FileItemViewModel>();

            var removedCount = 0;
            foreach (var item in Items.ToList())
            {
                if (!latestByPath.ContainsKey(item.Entry.FullPath))
                {
                    Items.Remove(item);
                    hasChanges = true;
                    removedCount++;

                    if (removedCount % 64 == 0)
                        await Task.Yield();
                }
            }

            var processedCount = 0;
            foreach (var (fullPath, latestEntry) in latestByPath)
            {
                if (existingByPath.TryGetValue(fullPath, out var existingItem))
                {
                    if (!AreEntriesEquivalent(existingItem.Entry, latestEntry))
                    {
                        PreserveComputedFields(existingItem.Entry, latestEntry);
                        existingItem.ApplyEntrySnapshot(latestEntry);
                        affectedItems.Add(existingItem);
                        reorderCandidates.Add(existingItem);
                        hasChanges = true;
                    }

                    processedCount++;
                    if (processedCount % 128 == 0)
                        await Task.Yield();

                    continue;
                }

                latestEntry.TypeDescription = GetTypeDescriptionFast(latestEntry);
                latestEntry.IconIndex = 0;
                var newItem = new FileItemViewModel(latestEntry);
                Items.Add(newItem);
                affectedItems.Add(newItem);
                 reorderCandidates.Add(newItem);
                hasChanges = true;

                processedCount++;
                if (processedCount % 128 == 0)
                    await Task.Yield();
            }

            if (reorderCandidates.Count > 0)
            {
                hasChanges |= await ApplyTargetedReorderAsync(reorderCandidates);
            }

            if (hasChanges && selectedPaths.Count > 0)
            {
                SelectedItems.Clear();
                foreach (var item in Items)
                {
                    if (selectedPaths.Contains(item.Entry.FullPath))
                    {
                        SelectedItems.Add(item);
                    }
                }
            }

            IsEmpty = Items.Count == 0;
            UpdateStatusText();
            UpdateFreeSpace();

            _lastLoadCompletedUtc = DateTimeOffset.UtcNow;
            System.Diagnostics.Debug.WriteLine($"[Refresh] TOTAL (snapshot+diff+UI): {refreshSw.ElapsedMilliseconds} ms  path={path}");

            if (_watcherHandle is null || IsAutoRefreshPaused)
            {
                _watcherHandle?.Dispose();
                StartWatcher(path);
            }

            if (affectedItems.Count > 0)
            {
                if (_cloudStatusService.IsSyncRootDetected)
                {
                    foreach (var item in affectedItems)
                    {
                        _ = RefreshSingleCloudStatusAsync(item);
                    }
                }

            }
        }
        catch (OperationCanceledException)
        {
            // Refresh superseded.
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to refresh directory incrementally: {Path}", path);
        }
    }

    private async Task<bool> ApplyTargetedReorderAsync(IReadOnlyCollection<FileItemViewModel> candidates)
    {
        if (candidates.Count == 0 || Items.Count <= 1)
            return false;

        var sortedItems = ApplySort(Items.ToList());
        var expectedIndexByPath = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < sortedItems.Count; index++)
        {
            var path = sortedItems[index].Entry.FullPath;
            if (!expectedIndexByPath.ContainsKey(path))
            {
                expectedIndexByPath[path] = index;
            }
        }

        var moveCount = 0;
        foreach (var item in candidates)
        {
            if (!expectedIndexByPath.TryGetValue(item.Entry.FullPath, out var expectedIndex))
                continue;

            var currentIndex = Items.IndexOf(item);
            if (currentIndex < 0 || currentIndex == expectedIndex)
                continue;

            Items.Move(currentIndex, expectedIndex);
            moveCount++;

            if (moveCount % 64 == 0)
                await Task.Yield();
        }

        return moveCount > 0;
    }

    private static bool AreEntriesEquivalent(FileSystemEntry existing, FileSystemEntry latest)
    {
        return existing.IsDirectory == latest.IsDirectory &&
               existing.Size == latest.Size &&
               existing.DateModified.UtcTicks == latest.DateModified.UtcTicks &&
               existing.Attributes == latest.Attributes &&
               existing.ReparseTag == latest.ReparseTag;
    }

    private static void PreserveComputedFields(FileSystemEntry source, FileSystemEntry target)
    {
        target.CloudStatus = source.CloudStatus;
        target.TypeDescription = source.TypeDescription;
        target.IconIndex = source.IconIndex;
    }

    private bool HasMaterialChanges(IReadOnlyDictionary<string, FileSystemEntry> latestByPath)
    {
        if (Items.Count != latestByPath.Count)
            return true;

        foreach (var item in Items)
        {
            if (!latestByPath.TryGetValue(item.Entry.FullPath, out var latest))
                return true;

            if (item.Entry.IsDirectory != latest.IsDirectory)
                return true;

            if (item.Entry.DateModified.UtcTicks != latest.DateModified.UtcTicks)
                return true;

            if (!item.Entry.IsDirectory && item.Entry.Size != latest.Size)
                return true;

            if (item.Entry.Attributes != latest.Attributes)
                return true;

            if (item.Entry.ReparseTag != latest.ReparseTag)
                return true;
        }

        return false;
    }

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

    private void ApplyFolderSortPreference(string folderPath)
    {
        var normalizedPath = NormalizeNavigationPath(folderPath);
        var persisted = _settingsService.GetFolderSort(normalizedPath);
        if (persisted is null)
            return;

        var persistedColumn = string.IsNullOrWhiteSpace(persisted.SortColumn)
            ? "Name"
            : persisted.SortColumn;

        SortColumn = persistedColumn;
        SortAscending = persisted.SortAscending;
        TabState.SortColumn = SortColumn;
        TabState.SortAscending = SortAscending;
    }

    private void PersistCurrentFolderSortPreference()
    {
        if (string.IsNullOrWhiteSpace(CurrentPath))
            return;

        try
        {
            _settingsService.SaveFolderSort(CurrentPath, new FolderSortSettings
            {
                SortColumn = SortColumn,
                SortAscending = SortAscending
            });

            _ = _settingsService.SaveAsync();
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to persist folder sort settings for {Path}", CurrentPath);
        }
    }

    private void StartWatcher(string path)
    {
        // Auto-refresh from file system notifications is intentionally disabled to avoid
        // frequent redraws; refresh is driven explicitly by UI actions and focus/tab events.
        _watcherHandle?.Dispose();
        _watcherHandle = null;
        IsAutoRefreshPaused = true;
    }

    private void DispatchFileSystemChange(FileSystemChangeEvent change)
    {
        void Execute()
        {
            try
            {
                HandleFileSystemChange(change);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed to process file system change {ChangeType} for {Path}", change.ChangeType, change.FullPath);
            }
        }

        if (_uiContext is null || ReferenceEquals(SynchronizationContext.Current, _uiContext))
        {
            Execute();
            return;
        }

        _uiContext.Post(_ => Execute(), null);
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
                        entry.TypeDescription = GetTypeDescriptionFast(entry);
                        var vm = new FileItemViewModel(entry) { IsNewItem = true };
                        Items.Add(vm);

                        if (_cloudStatusService.IsSyncRootDetected)
                        {
                            _ = RefreshSingleCloudStatusAsync(vm);
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
                    var previousPath = renamed.Entry.FullPath;
                    renamed.Entry.Name = change.Name;
                    renamed.Entry.FullPath = change.FullPath;
                    renamed.Entry.Extension = Path.GetExtension(change.FullPath);
                    renamed.Entry.TypeDescription = GetTypeDescriptionFast(renamed.Entry);
                    renamed.NotifyNameChanged();
                    renamed.NotifyRenamed();

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



















    private static string GetTypeDescriptionFast(FileSystemEntry entry)
    {
        if (entry.IsDirectory)
            return "File folder";

        if (string.IsNullOrWhiteSpace(entry.Extension))
            return "File";

        return $"{entry.Extension.TrimStart('.').ToUpperInvariant()} File";
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
        _ = _tabPathStateService.DeleteByTabIdAsync(TabState.Id);
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

    /// <summary>Applies the latest file system snapshot and raises dependent property updates.</summary>
    public void ApplyEntrySnapshot(FileSystemEntry latest)
    {
        Entry.Name = latest.Name;
        Entry.FullPath = latest.FullPath;
        Entry.IsDirectory = latest.IsDirectory;
        Entry.Size = latest.Size;
        Entry.DateModified = latest.DateModified;
        Entry.TypeDescription = latest.TypeDescription;
        Entry.Extension = latest.Extension;
        Entry.Attributes = latest.Attributes;
        Entry.ReparseTag = latest.ReparseTag;
        Entry.IconIndex = latest.IconIndex;

        OnPropertyChanged(nameof(Name));
        OnPropertyChanged(nameof(FullPath));
        OnPropertyChanged(nameof(IsDirectory));
        OnPropertyChanged(nameof(SizeText));
        OnPropertyChanged(nameof(DateModifiedText));
        OnPropertyChanged(nameof(TypeDescription));
        OnPropertyChanged(nameof(Extension));
        OnPropertyChanged(nameof(IconGlyph));
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
