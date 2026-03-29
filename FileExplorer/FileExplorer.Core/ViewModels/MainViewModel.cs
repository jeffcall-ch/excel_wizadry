// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Core.ViewModels;

/// <summary>
/// Root ViewModel managing tabs, the tree view, command bar state, and global operations.
/// Bound to <c>MainWindow.xaml</c>.
/// </summary>
public sealed partial class MainViewModel : ObservableObject, IDisposable
{
    private const string DefaultStartupPath = @"C:\Users\szil\Shortcuts_to_folders";

    private readonly IFileSystemService _fileSystemService;
    private readonly IShellIntegrationService _shellService;
    private readonly ICloudStatusService _cloudStatusService;
    private readonly ISettingsService _settingsService;
    private readonly IPreviewService _previewService;
    private readonly ISearchService _searchService;
    private readonly ILogger<MainViewModel> _logger;
    private readonly Func<TabState?, TabContentViewModel> _tabFactory;
    private CancellationTokenSource? _searchCts;

    /// <summary>All open tabs.</summary>
    public ObservableCollection<TabContentViewModel> Tabs { get; } = [];

    /// <summary>Tree view model for the left pane.</summary>
    public TreeViewModel TreeView { get; }

    [ObservableProperty]
    private TabContentViewModel? _activeTab;

    [ObservableProperty]
    private bool _showHiddenFiles;

    [ObservableProperty]
    private bool _showFileExtensions = true;

    [ObservableProperty]
    private bool _showPreviewPane;

    [ObservableProperty]
    private double _navigationPaneWidth = 250;

    [ObservableProperty]
    private double _previewPaneWidth = 300;

    [ObservableProperty]
    private bool _hasClipboardContent;

    [ObservableProperty]
    private FileOperationProgress? _currentOperationProgress;

    /// <summary>Whether the active tab can navigate back.</summary>
    public bool CanGoBack => ActiveTab?.History.CanGoBack ?? false;

    /// <summary>Whether the active tab can navigate forward.</summary>
    public bool CanGoForward => ActiveTab?.History.CanGoForward ?? false;

    /// <summary>Whether the active tab has a parent directory.</summary>
    public bool CanGoUp => ActiveTab is not null && Path.GetDirectoryName(ActiveTab.CurrentPath) is not null;

    /// <summary>Whether there are items selected in the active tab.</summary>
    public bool HasSelection => ActiveTab?.SelectedItems.Count > 0;

    /// <summary>Whether undo is available.</summary>
    public bool CanUndo => _fileSystemService.CanUndo;

    /// <summary>OneDrive column should be visible.</summary>
    public bool ShowOneDriveColumn => _cloudStatusService.IsSyncRootDetected;

    /// <summary>Initializes a new instance of the <see cref="MainViewModel"/> class.</summary>
    public MainViewModel(
        IFileSystemService fileSystemService,
        IShellIntegrationService shellService,
        ICloudStatusService cloudStatusService,
        ISettingsService settingsService,
        IPreviewService previewService,
        ISearchService searchService,
        TreeViewModel treeViewModel,
        Func<TabState?, TabContentViewModel> tabFactory,
        ILogger<MainViewModel> logger)
    {
        _fileSystemService = fileSystemService;
        _shellService = shellService;
        _cloudStatusService = cloudStatusService;
        _settingsService = settingsService;
        _previewService = previewService;
        _searchService = searchService;
        _logger = logger;
        _tabFactory = tabFactory;
        TreeView = treeViewModel;

        TreeView.NavigationRequested += OnTreeNavigationRequested;
    }

    /// <summary>Initializes the application: loads settings, detects OneDrive, and opens the initial tab.</summary>
    public async Task InitializeAsync(string? startupPath = null)
    {
        await _settingsService.LoadAsync();
        var settings = _settingsService.Settings;
        ShowHiddenFiles = settings.ShowHiddenFiles;
        ShowFileExtensions = settings.ShowFileExtensions;
        ShowPreviewPane = settings.ShowPreviewPane;
        NavigationPaneWidth = settings.NavigationPaneWidth;
        PreviewPaneWidth = settings.PreviewPaneWidth;

        // Detect OneDrive sync roots
        await _cloudStatusService.DetectSyncRootsAsync();
        OnPropertyChanged(nameof(ShowOneDriveColumn));

        // Load tree view
        await TreeView.LoadAsync();

        var initialPath = ResolveInitialStartupPath(startupPath, settings.NewTabDefaultPath);
        await OpenTabAsync(initialPath, null);
    }

    #region Tab Management

    /// <summary>Opens a new tab.</summary>
    [RelayCommand]
    public async Task NewTabAsync()
    {
        var defaultPath = _settingsService.Settings.NewTabDefaultPath;

        if (!string.IsNullOrEmpty(defaultPath) && Directory.Exists(defaultPath))
        {
            await OpenTabAsync(defaultPath, null);
        }
        else
        {
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            await OpenTabAsync(userProfile, null);
        }
    }

    public async Task OpenFolderInNewTabAsync(string path)
    {
        var normalizedPath = NormalizeStartupPath(path);
        if (!Directory.Exists(normalizedPath))
            return;

        await OpenTabAsync(normalizedPath, null);
    }

    private async Task OpenTabAsync(string path, TabState? state)
    {
        var tab = _tabFactory(state);
        tab.Navigated += OnTabNavigated;
        Tabs.Add(tab);
        ActiveTab = tab;
        await tab.NavigateAsync(path);
    }

    private string ResolveInitialStartupPath(string? startupPath, string? fallbackNewTabPath)
    {
        if (!string.IsNullOrWhiteSpace(startupPath))
        {
            var normalizedStartupPath = NormalizeStartupPath(startupPath);

            if (Directory.Exists(normalizedStartupPath))
                return normalizedStartupPath;

            if (File.Exists(normalizedStartupPath))
            {
                var parent = Path.GetDirectoryName(normalizedStartupPath);
                if (!string.IsNullOrEmpty(parent) && Directory.Exists(parent))
                    return parent;
            }

            _logger.LogWarning("Startup path does not exist: {Path}", startupPath);
        }

        if (Directory.Exists(DefaultStartupPath))
            return DefaultStartupPath;

        if (!string.IsNullOrWhiteSpace(fallbackNewTabPath) && Directory.Exists(fallbackNewTabPath))
            return fallbackNewTabPath;

        return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    }

    private static string NormalizeStartupPath(string path)
    {
        return TabContentViewModel.NormalizeNavigationPath(path);
    }

    /// <summary>Closes the specified tab.</summary>
    [RelayCommand]
    public void CloseTab(TabContentViewModel tab)
    {
        if (tab.IsPinned) return;
        if (Tabs.Count <= 1)
        {
            // Don't close the last tab — navigate to default instead
            return;
        }

        var index = Tabs.IndexOf(tab);
        tab.Navigated -= OnTabNavigated;
        tab.Dispose();
        Tabs.Remove(tab);

        if (ActiveTab == tab)
        {
            ActiveTab = Tabs[Math.Min(index, Tabs.Count - 1)];
        }
    }

    /// <summary>Closes all tabs except the specified one.</summary>
    [RelayCommand]
    public void CloseOtherTabs(TabContentViewModel keepTab)
    {
        var toClose = Tabs.Where(t => t != keepTab && !t.IsPinned).ToList();
        foreach (var tab in toClose)
        {
            tab.Navigated -= OnTabNavigated;
            tab.Dispose();
            Tabs.Remove(tab);
        }
        ActiveTab = keepTab;
    }

    /// <summary>Closes all tabs to the right of the specified tab.</summary>
    [RelayCommand]
    public void CloseTabsToRight(TabContentViewModel referenceTab)
    {
        var index = Tabs.IndexOf(referenceTab);
        var toClose = Tabs.Skip(index + 1).Where(t => !t.IsPinned).ToList();
        foreach (var tab in toClose)
        {
            tab.Navigated -= OnTabNavigated;
            tab.Dispose();
            Tabs.Remove(tab);
        }
    }

    /// <summary>Duplicates the specified tab.</summary>
    [RelayCommand]
    public async Task DuplicateTabAsync(TabContentViewModel sourceTab)
    {
        var tab = _tabFactory(null);
        tab.Navigated += OnTabNavigated;
        var index = Tabs.IndexOf(sourceTab);
        Tabs.Insert(index + 1, tab);
        ActiveTab = tab;
        await tab.NavigateAsync(sourceTab.CurrentPath);
    }

    /// <summary>Toggles the pin state of a tab.</summary>
    [RelayCommand]
    public void TogglePinTab(TabContentViewModel tab)
    {
        tab.IsPinned = !tab.IsPinned;
        tab.TabState.IsPinned = tab.IsPinned;
    }

    /// <summary>Navigates to a specific tab by index (1-based for Ctrl+1-8).</summary>
    [RelayCommand]
    public void GoToTab(int index)
    {
        if (index == 9)
        {
            // Ctrl+9 = last tab
            ActiveTab = Tabs.LastOrDefault();
            return;
        }

        // Ctrl+1-8 = 0-7 indexed
        var zeroIndex = index - 1;
        if (zeroIndex >= 0 && zeroIndex < Tabs.Count)
        {
            ActiveTab = Tabs[zeroIndex];
        }
    }

    /// <summary>Cycles to the next tab.</summary>
    [RelayCommand]
    public void NextTab()
    {
        if (ActiveTab is null || Tabs.Count <= 1) return;
        var index = Tabs.IndexOf(ActiveTab);
        ActiveTab = Tabs[(index + 1) % Tabs.Count];
    }

    /// <summary>Cycles to the previous tab.</summary>
    [RelayCommand]
    public void PreviousTab()
    {
        if (ActiveTab is null || Tabs.Count <= 1) return;
        var index = Tabs.IndexOf(ActiveTab);
        ActiveTab = Tabs[(index - 1 + Tabs.Count) % Tabs.Count];
    }

    #endregion

    #region Navigation Commands

    /// <summary>Navigates back in the active tab.</summary>
    [RelayCommand]
    public async Task GoBackAsync()
    {
        if (ActiveTab is not null) await ActiveTab.GoBackAsync();
        NotifyNavigationState();
    }

    /// <summary>Navigates forward in the active tab.</summary>
    [RelayCommand]
    public async Task GoForwardAsync()
    {
        if (ActiveTab is not null) await ActiveTab.GoForwardAsync();
        NotifyNavigationState();
    }

    /// <summary>Navigates up in the active tab.</summary>
    [RelayCommand]
    public async Task GoUpAsync()
    {
        if (ActiveTab is not null) await ActiveTab.GoUpAsync();
        NotifyNavigationState();
    }

    /// <summary>Refreshes the active tab.</summary>
    [RelayCommand]
    public async Task RefreshAsync()
    {
        if (ActiveTab is not null) await ActiveTab.RefreshAsync();
    }

    /// <summary>Navigates to a path typed in the address bar.</summary>
    [RelayCommand]
    public async Task NavigateToAddressAsync(string address)
    {
        if (ActiveTab is null) return;
        await ActiveTab.NavigateAsync(address);
        NotifyNavigationState();
    }

    /// <summary>Searches for files matching the query in the current directory using rq engine.</summary>
    public async Task SearchAsync(string query)
    {
        if (ActiveTab is null || string.IsNullOrWhiteSpace(query)) return;

        // Cancel any previous search
        _searchCts?.Cancel();
        _searchCts = new CancellationTokenSource();
        var ct = _searchCts.Token;

        var searchDir = ActiveTab.CurrentPath;
        _logger.LogInformation("Search: '{Query}' in {Dir}", query, searchDir);

        ActiveTab.IsLoading = true;
        ActiveTab.Items.Clear();

        try
        {
            var paths = await _searchService.SearchAsync(searchDir, query, ct);
            ct.ThrowIfCancellationRequested();

            foreach (var path in paths)
            {
                try
                {
                    FileSystemEntry entry;
                    if (Directory.Exists(path))
                    {
                        var di = new DirectoryInfo(path);
                        entry = new FileSystemEntry
                        {
                            FullPath = di.FullName,
                            Name = di.Name,
                            IsDirectory = true,
                            DateModified = di.LastWriteTime,
                            Attributes = di.Attributes,
                            Extension = string.Empty,
                            TypeDescription = "File folder"
                        };
                    }
                    else if (File.Exists(path))
                    {
                        var fi = new FileInfo(path);
                        entry = new FileSystemEntry
                        {
                            FullPath = fi.FullName,
                            Name = fi.Name,
                            IsDirectory = false,
                            Size = fi.Length,
                            DateModified = fi.LastWriteTime,
                            Attributes = fi.Attributes,
                            Extension = fi.Extension,
                            TypeDescription = string.IsNullOrEmpty(fi.Extension) ? "File" : $"{fi.Extension.TrimStart('.').ToUpperInvariant()} File"
                        };
                    }
                    else continue;

                    ActiveTab.Items.Add(new FileItemViewModel(entry));
                }
                catch
                {
                    // Skip inaccessible files
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Search was superseded
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Search failed");
        }
        finally
        {
            ActiveTab.IsLoading = false;
        }
    }

    /// <summary>Clears search and returns to the current directory listing.</summary>
    public async Task ClearSearchAsync()
    {
        _searchCts?.Cancel();
        if (ActiveTab is not null)
            await ActiveTab.NavigateAsync(ActiveTab.CurrentPath);
    }

    #endregion

    #region File Operations

    /// <summary>Undoes the last file operation.</summary>
    [RelayCommand]
    public async Task UndoAsync()
    {
        await _fileSystemService.UndoLastOperationAsync();
        await RefreshAsync();
        OnPropertyChanged(nameof(CanUndo));
    }

    /// <summary>Opens properties for selected items.</summary>
    [RelayCommand]
    public void ShowProperties(nint hwnd)
    {
        if (ActiveTab is null) return;
        var paths = ActiveTab.SelectedItems.Select(i => i.Entry.FullPath).ToList();
        if (paths.Count > 0)
        {
            _shellService.ShowProperties(hwnd, paths);
        }
    }

    #endregion

    #region Settings

    /// <summary>Toggles visibility of hidden files.</summary>
    [RelayCommand]
    public void ToggleHiddenFiles()
    {
        ShowHiddenFiles = !ShowHiddenFiles;
        _settingsService.Settings.ShowHiddenFiles = ShowHiddenFiles;
    }

    /// <summary>Toggles visibility of file extensions.</summary>
    [RelayCommand]
    public void ToggleFileExtensions()
    {
        ShowFileExtensions = !ShowFileExtensions;
        _settingsService.Settings.ShowFileExtensions = ShowFileExtensions;
    }

    /// <summary>Toggles the preview pane.</summary>
    [RelayCommand]
    public void TogglePreviewPane()
    {
        ShowPreviewPane = !ShowPreviewPane;
        _settingsService.Settings.ShowPreviewPane = ShowPreviewPane;
    }

    #endregion

    #region Session Save

    /// <summary>Saves the current session state (tabs, settings).</summary>
    public async Task SaveSessionAsync()
    {
        _settingsService.Settings.NavigationPaneWidth = NavigationPaneWidth;
        _settingsService.Settings.PreviewPaneWidth = PreviewPaneWidth;

        await _settingsService.SaveTabSessionAsync(Tabs.Select(t => t.TabState).ToList());
        await _settingsService.SaveAsync();
        _logger.LogInformation("Session saved: {Count} tabs", Tabs.Count);
    }

    #endregion

    #region Event Handlers

    partial void OnActiveTabChanged(TabContentViewModel? value)
    {
        NotifyNavigationState();
        OnPropertyChanged(nameof(HasSelection));
    }

    private void OnTreeNavigationRequested(object? sender, string path)
    {
        if (ActiveTab is not null)
        {
            _ = ActiveTab.NavigateAsync(path);
        }
    }

    private void OnTabNavigated(object? sender, string path)
    {
        NotifyNavigationState();
    }

    private void NotifyNavigationState()
    {
        OnPropertyChanged(nameof(CanGoBack));
        OnPropertyChanged(nameof(CanGoForward));
        OnPropertyChanged(nameof(CanGoUp));
    }

    #endregion

    /// <inheritdoc/>
    public void Dispose()
    {
        TreeView.NavigationRequested -= OnTreeNavigationRequested;
        foreach (var tab in Tabs)
        {
            tab.Navigated -= OnTabNavigated;
            tab.Dispose();
        }
        TreeView.Dispose();
    }
}
