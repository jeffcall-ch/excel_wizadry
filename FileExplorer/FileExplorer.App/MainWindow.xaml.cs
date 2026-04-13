// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Diagnostics;
using System.Collections.Concurrent;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using FileExplorer.Core.ViewModels;
using FileExplorer.Core.Models;
using FileExplorer.Core.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation;
using Windows.Storage;
using WinRT.Interop;

namespace FileExplorer.App;

/// <summary>
/// Main application window hosting tabs, tree view, file list, and command bar.
/// </summary>
public sealed partial class MainWindow : Window
{
    private const string BaseWindowTitle = "File Explorer";
    private const int IDC_ARROW = 32512;
    private const int IDC_SIZEWE = 32644;
    private const int IMAGE_ICON = 1;
    private const uint LR_LOADFROMFILE = 0x00000010;
    private const uint LR_DEFAULTSIZE = 0x00000040;
    private const uint WM_SETICON = 0x0080;
    private static readonly nint ICON_SMALL = new(0);
    private static readonly nint ICON_BIG = new(1);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern nint LoadCursor(nint hInstance, int lpCursorName);

    [DllImport("user32.dll")]
    private static extern nint SetCursor(nint hCursor);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern nint LoadImage(nint hInst, string name, uint type, int cx, int cy, uint fuLoad);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern nint SendMessage(nint hWnd, uint msg, nint wParam, nint lParam);

    private readonly DispatcherQueue _dispatcherQueue;
    private readonly IFileSystemService _fileSystemService;
    private readonly ISettingsService _settingsService;
    private readonly IDirectorySnapshotCache _snapshotCache;
    private readonly HashSet<TabContentViewModel> _tabPropertySubscriptions = [];
    private readonly Dictionary<Guid, TabViewItem> _tabItemsById = [];
    private bool _suppressTabSelectionChanged;
    private bool _windowWasDeactivated;
    private TreeNodeViewModel? _pendingSelectedTreeNode;
    private string? _lastSyncedTreePath;
    private readonly string _shortcutsFolderPath;
    private bool _isFileInUseWarningOpen;
    private readonly object _shortcutCacheLock = new();
    private readonly Dictionary<string, ShortcutFolderCacheData> _shortcutFolderCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _shortcutPrefetchInProgress = new(StringComparer.OrdinalIgnoreCase);
    private List<ShortcutRootEntryData> _shortcutRootCache = [];
    private readonly ConcurrentDictionary<string, IReadOnlyList<string>> _breadcrumbSiblingCache = new(StringComparer.OrdinalIgnoreCase);
    private MenuFlyout? _shortcutLauncherFlyout;
    private readonly Dictionary<string, MenuFlyout> _hubChildrenFlyoutCache = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _breadcrumbSiblingCts;
    private int _breadcrumbSiblingRequestId;
    private bool _shortcutPreloadStarted;
    private bool _shortcutDeepPreloadStarted;
    private bool _snapshotPrefetchStarted;
    private bool _isShortcutOpenInProgress;
    private bool _startupWindowStateApplied;
    private bool _isActivationRefreshInProgress;
    private DateTimeOffset _lastActivationRefreshUtc = DateTimeOffset.MinValue;
    private const int ShortcutMenuMaxItems = 50;
    private const int ShortcutMenuMaxSubfolders = 50;
    private const int ShortcutMenuMaxDepth = 64;
    private static readonly TimeSpan ActivationRefreshMinInterval = TimeSpan.FromSeconds(1);
    private const double FlyoutDetailsFontSize = 12;
    private const double FlyoutDetailsMinHeight = 22;

    public sealed record ShortcutHubRootItem(string DisplayName, string? FolderPath, string? ShortcutPath, bool IsFolder);

    private sealed record ShortcutRootEntryData(string DisplayName, string? FolderPath, string? ShortcutPath);

    private sealed record ShortcutSubfolderData(string DisplayName, string FolderPath);

    private sealed record ShortcutFileData(string Name, string FullPath);

    private sealed record ShortcutFolderCacheData(
        IReadOnlyList<ShortcutSubfolderData> Subfolders,
        IReadOnlyList<ShortcutFileData> Files,
        bool IsFullyLoaded);

    private sealed class ShortcutFolderMenuContext
    {
        public ShortcutFolderMenuContext(string folderPath, int depth)
        {
            FolderPath = folderPath;
            Depth = depth;
        }

        public string FolderPath { get; }

        public int Depth { get; }

        public bool IsLoaded { get; set; }

        public bool IsLoading { get; set; }
    }

    private static MenuFlyoutItem CreateDetailsViewFlyoutItem(string text, string? iconGlyph = null, bool isEnabled = true)
    {
        var item = new MenuFlyoutItem
        {
            Text = text,
            IsEnabled = isEnabled
        };

        if (!string.IsNullOrWhiteSpace(iconGlyph))
        {
            item.Icon = new FontIcon { Glyph = iconGlyph };
        }

        ApplyDetailsViewFlyoutItemStyle(item);
        return item;
    }

    private static void ApplyDetailsViewFlyoutItemStyle(MenuFlyoutItem item)
    {
        item.FontSize = FlyoutDetailsFontSize;
        item.MinHeight = FlyoutDetailsMinHeight;
        item.Padding = new Thickness(8, 0, 8, 0);
        item.VerticalContentAlignment = VerticalAlignment.Center;
    }

    private static void ApplyDetailsViewFlyoutSubItemStyle(MenuFlyoutSubItem item)
    {
        item.FontSize = FlyoutDetailsFontSize;
        item.MinHeight = FlyoutDetailsMinHeight;
        item.Padding = new Thickness(8, 0, 8, 0);
        item.VerticalContentAlignment = VerticalAlignment.Center;
    }

    /// <summary>Gets the main view model.</summary>
    public MainViewModel ViewModel { get; }

    /// <summary>Gets the native window handle.</summary>
    public nint Hwnd { get; }

    /// <summary>Initializes a new instance of the <see cref="MainWindow"/> class.</summary>
    public MainWindow()
    {
        InitializeComponent();

        TabViewControl.AddTabButtonClick += TabView_AddTabButtonClick;
        TabViewControl.TabCloseRequested += TabView_TabCloseRequested;
        TabViewControl.SelectionChanged += TabView_SelectionChanged;
        TabViewControl.TabItemsChanged += TabView_TabItemsChanged;

        Title = BaseWindowTitle;
        SystemBackdrop = new MicaBackdrop();

        Hwnd = WindowNative.GetWindowHandle(this);
        _dispatcherQueue = DispatcherQueue.GetForCurrentThread();
        _fileSystemService = App.Services.GetRequiredService<IFileSystemService>();
        _settingsService = App.Services.GetRequiredService<ISettingsService>();
        _snapshotCache = App.Services.GetRequiredService<IDirectorySnapshotCache>();
        ViewModel = App.Services.GetRequiredService<MainViewModel>();
        ViewModel.Tabs.CollectionChanged += Tabs_CollectionChanged;
        // Single authoritative ActivateTab path: whenever MainViewModel.ActiveTab changes,
        // bind the FileList and Inspector to the new tab. This covers new-tab, duplicate-tab,
        // open-folder-in-new-tab, and Ctrl+1-9 keyboard shortcuts — all paths that change
        // ActiveTab without going through TabView_SelectionChanged.
        ViewModel.PropertyChanged += ViewModel_ActiveTabChanged;
        RootGrid.DataContext = ViewModel;
        _shortcutsFolderPath = ResolveShortcutsFolderPath();
        _fileSystemService.OperationError += FileSystemService_OperationError;

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(TabViewControl);
        ApplyWindowIcon();
        Activated += MainWindow_Activated;

        RegisterKeyboardAccelerators();
        Closed += MainWindow_Closed;
        NavigationTree.Loaded += NavigationTree_Loaded;

        StartShortcutCachePreload();

        _ = InitializeAsync();
    }

    private void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState == WindowActivationState.Deactivated)
        {
            _windowWasDeactivated = true;
            return;
        }

        EnsureStartupMaximized();

        if (_windowWasDeactivated)
        {
            _windowWasDeactivated = false;
            _ = RefreshAfterWindowActivationAsync();
        }
    }

    private async Task RefreshAfterWindowActivationAsync()
    {
        if (_isActivationRefreshInProgress || ViewModel.ActiveTab is null)
            return;

        var now = DateTimeOffset.UtcNow;
        if (now - _lastActivationRefreshUtc < ActivationRefreshMinInterval)
            return;

        // Capture the tab NOW (synchronously, before any await) so that a concurrent
        // tab-switch (TabView_SelectionChanged fires before Activated on pointer-click)
        // cannot redirect this scan to the newly-selected tab.
        var tabToRefresh = ViewModel.ActiveTab;

        _isActivationRefreshInProgress = true;
        _lastActivationRefreshUtc = now;

        try
        {
            if (!string.IsNullOrWhiteSpace(tabToRefresh.CurrentPath))
            {
                System.Diagnostics.Debug.WriteLine($"[WindowActivation] refresh start  path={tabToRefresh.CurrentPath}");
                await tabToRefresh.RefreshAsync();
                System.Diagnostics.Debug.WriteLine($"[WindowActivation] refresh done   path={tabToRefresh.CurrentPath}");
                UpdateStatusBar();
            }
        }
        finally
        {
            _isActivationRefreshInProgress = false;
        }
    }

    private void EnsureStartupMaximized()
    {
        if (_startupWindowStateApplied)
            return;

        _startupWindowStateApplied = true;

        try
        {
            if (AppWindow.Presenter is OverlappedPresenter overlappedPresenter)
            {
                if (overlappedPresenter.State != OverlappedPresenterState.Maximized)
                    overlappedPresenter.Maximize();

                return;
            }

            AppWindow.SetPresenter(AppWindowPresenterKind.Overlapped);
            if (AppWindow.Presenter is OverlappedPresenter presenter)
                presenter.Maximize();
        }
        catch
        {
            // Window state setup should not block startup.
        }
    }

    private void ApplyWindowIcon()
    {
        try
        {
            var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "LSApp.ico");
            if (!File.Exists(iconPath))
                return;

            AppWindow.SetIcon(iconPath);

            var hIcon = LoadImage(nint.Zero, iconPath, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE);
            if (hIcon != nint.Zero)
            {
                SendMessage(Hwnd, WM_SETICON, ICON_SMALL, hIcon);
                SendMessage(Hwnd, WM_SETICON, ICON_BIG, hIcon);
            }
        }
        catch
        {
            // Icon setup should never block app startup.
        }
    }

    private async Task InitializeAsync()
    {
        try
        {
            await ViewModel.InitializeAsync(App.StartupPath);
            SyncTabPropertySubscriptions();

            UpdateTabBindings();
            UpdateStatusBar();

            TreeColumn.Width = new GridLength(0);
            PreviewColumn.Width = new GridLength(ViewModel.PreviewPaneWidth > 0 ? ViewModel.PreviewPaneWidth : 320);

            if (ViewModel.ActiveTab is not null)
            {
                SubscribeToTabNavigation(ViewModel.ActiveTab);
                FileList.BindToTab(ViewModel.ActiveTab);
                InspectorSidebar.BindToTab(ViewModel.ActiveTab);
                UpdateWindowTitle(ViewModel.ActiveTab);
                UpdateAddressBar(ViewModel.ActiveTab.CurrentPath);
                UpdateSearchPlaceholder(ViewModel.ActiveTab.CurrentPath);
            }

            StartShortcutCachePreload();
        }
        catch (Exception ex)
        {
            ShowError($"Initialization failed: {ex.Message}");
        }
    }

    #region Tab Events

    private void TabView_AddTabButtonClick(TabView sender, object args)
    {
        System.Diagnostics.Debug.WriteLine("[NewTab] button clicked");
        _ = ViewModel.NewTabAsync();
        UpdateTabBindings();
    }

    private void TabView_TabCloseRequested(TabView sender, TabViewTabCloseRequestedEventArgs args)
    {
        if (args.Item is TabViewItem tabViewItem && tabViewItem.DataContext is TabContentViewModel tab)
        {
            ViewModel.CloseTab(tab);
            UpdateTabBindings();
            return;
        }

        if (args.Item is TabContentViewModel fallbackTab)
        {
            ViewModel.CloseTab(fallbackTab);
            UpdateTabBindings();
        }
    }

    private void TabView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressTabSelectionChanged)
            return;

        if (TabViewControl.SelectedItem is TabViewItem tabViewItem && tabViewItem.DataContext is TabContentViewModel tab)
        {
            ActivateTab(tab, syncTabViewSelection: false);
            ActivateRightPaneForKeyboardNavigation();

            return;
        }

        if (TabViewControl.SelectedItem is TabContentViewModel fallbackTab)
        {
            ActivateTab(fallbackTab, syncTabViewSelection: false);
            ActivateRightPaneForKeyboardNavigation();
        }
    }

    private void UpdateTabBindings()
    {
        var orderedTabs = ViewModel.Tabs.ToList();
        var targetIds = orderedTabs.Select(tab => tab.TabState.Id).ToHashSet();

        foreach (var staleEntry in _tabItemsById.ToList())
        {
            if (!targetIds.Contains(staleEntry.Key))
            {
                TabViewControl.TabItems.Remove(staleEntry.Value);
                _tabItemsById.Remove(staleEntry.Key);
            }
        }

        for (var index = 0; index < orderedTabs.Count; index++)
        {
            var tab = orderedTabs[index];
            if (!_tabItemsById.TryGetValue(tab.TabState.Id, out var tabViewItem))
            {
                tabViewItem = new TabViewItem
                {
                    DataContext = tab,
                    IsClosable = !tab.IsPinned
                };

                _tabItemsById[tab.TabState.Id] = tabViewItem;
                TabViewControl.TabItems.Insert(index, tabViewItem);
            }
            else
            {
                var currentIndex = TabViewControl.TabItems.IndexOf(tabViewItem);
                if (currentIndex >= 0 && currentIndex != index)
                {
                    TabViewControl.TabItems.RemoveAt(currentIndex);
                    TabViewControl.TabItems.Insert(index, tabViewItem);
                }
            }

            RefreshTabHeader(tab);
        }

        if (ViewModel.ActiveTab is not null && _tabItemsById.TryGetValue(ViewModel.ActiveTab.TabState.Id, out var selectedTabViewItem))
        {
            if (!ReferenceEquals(TabViewControl.SelectedItem, selectedTabViewItem))
            {
                _suppressTabSelectionChanged = true;
                try
                {
                    TabViewControl.SelectedItem = selectedTabViewItem;
                }
                finally
                {
                    _suppressTabSelectionChanged = false;
                }
            }
        }

        _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, RefreshVisibleTabHeaders);
    }

    private void TabView_TabItemsChanged(TabView sender, object args)
    {
        _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, UpdateTabBindings);
    }

    private void Tabs_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<TabContentViewModel>())
            {
                RemoveTabPropertySubscription(item);
            }
        }

        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<TabContentViewModel>())
            {
                EnsureTabPropertySubscription(item);
            }
        }

        _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, RefreshVisibleTabHeaders);
    }

    private void SyncTabPropertySubscriptions()
    {
        foreach (var tab in ViewModel.Tabs)
        {
            EnsureTabPropertySubscription(tab);
        }

        foreach (var staleTab in _tabPropertySubscriptions.ToList())
        {
            if (!ViewModel.Tabs.Contains(staleTab))
            {
                RemoveTabPropertySubscription(staleTab);
            }
        }
    }

    private void ViewModel_ActiveTabChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (!string.Equals(e.PropertyName, nameof(MainViewModel.ActiveTab), StringComparison.Ordinal))
            return;

        var tab = ViewModel.ActiveTab;
        if (tab is null)
            return;

        // Always marshal to the UI thread — PropertyChanged can fire from any context.
        _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Normal, () =>
        {
            // ActivateTab is idempotent for the same tab (BindToTab guards with ReferenceEquals).
            // TabView_SelectionChanged may also call ActivateTab for the same tab — that is fine.
            ActivateTab(tab, syncTabViewSelection: true);
        });
    }

    private void EnsureTabPropertySubscription(TabContentViewModel tab)
    {
        if (!_tabPropertySubscriptions.Add(tab))
            return;

        tab.PropertyChanged += OnTabPropertyChanged;
    }

    private void RemoveTabPropertySubscription(TabContentViewModel tab)
    {
        if (!_tabPropertySubscriptions.Remove(tab))
            return;

        tab.PropertyChanged -= OnTabPropertyChanged;
    }

    private void ActivateTab(TabContentViewModel tab, bool syncTabViewSelection)
    {
        var wasDifferentTab = !ReferenceEquals(ViewModel.ActiveTab, tab);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] ActiveTab set: 0 ms  (wasDifferent={wasDifferentTab}  items={tab.Items.Count}  path={tab.CurrentPath})");

        if (syncTabViewSelection && _tabItemsById.TryGetValue(tab.TabState.Id, out var tabViewItem) && !ReferenceEquals(TabViewControl.SelectedItem, tabViewItem))
        {
            _suppressTabSelectionChanged = true;
            try
            {
                TabViewControl.SelectedItem = tabViewItem;
            }
            finally
            {
                _suppressTabSelectionChanged = false;
            }
        }

        if (!ReferenceEquals(ViewModel.ActiveTab, tab))
        {
            ViewModel.ActiveTab = tab;
        }
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] ActiveTab set: {sw.ElapsedMilliseconds} ms");

        SubscribeToTabNavigation(tab);
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] SubscribeToTabNavigation: {sw.ElapsedMilliseconds} ms");

        FileList.BindToTab(tab);
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] BindToTab: {sw.ElapsedMilliseconds} ms  ({tab.Items.Count} items, sort={tab.SortColumn})");

        InspectorSidebar.BindToTab(tab);
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] InspectorSidebar.BindToTab: {sw.ElapsedMilliseconds} ms");

        UpdateWindowTitle(tab);
        UpdateAddressBar(tab.CurrentPath);
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] UpdateAddressBar: {sw.ElapsedMilliseconds} ms");

        UpdateStatusBar();
        UpdateSearchPlaceholder(tab.CurrentPath);
        System.Diagnostics.Debug.WriteLine($"[TabSwitch] TOTAL synchronous work: {sw.ElapsedMilliseconds} ms  path={tab.CurrentPath}");

        if (wasDifferentTab)
        {
            // Tab switches never trigger a background disk scan.
            // Cached items are shown instantly. The folder is re-scanned only when:
            //   - the user presses the Refresh button ( RefreshButton_Click )
            //   - the user returns to the app after being away ( RefreshAfterWindowActivationAsync )
            System.Diagnostics.Debug.WriteLine($"[TabSwitch] no refresh on tab switch → cached view  path={tab.CurrentPath}");
        }
    }

    #endregion

    #region Navigation

    private void RootGrid_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (sender is not UIElement element)
            return;

        var props = e.GetCurrentPoint(element).Properties;
        if (props.IsXButton1Pressed)
        {
            if (ViewModel.CanGoBack)
                _ = ViewModel.GoBackAsync();

            e.Handled = true;
            return;
        }

        if (props.IsXButton2Pressed)
        {
            if (ViewModel.CanGoForward)
                _ = ViewModel.GoForwardAsync();

            e.Handled = true;
            return;
        }

        var source = e.OriginalSource as DependencyObject;
        if (source is null)
            return;

        if (IsWithinNavigationPane(source))
            return;

        if (IsWithinTextInput(source))
            return;

        if (IsWithinFileListArea(source))
        {
            ActivateRightPaneForKeyboardNavigation();
        }
    }

    private void ActivateRightPaneForKeyboardNavigation()
    {
        _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, () =>
        {
            FileList.FocusForKeyboardNavigation();
        });
    }

    private void CloseShortcutLauncherFlyoutAndFocusRightPane()
    {
        _shortcutLauncherFlyout?.Hide();
        // Also close any open hub-children flyouts (shown from the inspector sidebar list).
        foreach (var flyout in _hubChildrenFlyoutCache.Values)
            flyout.Hide();
        ActivateRightPaneForKeyboardNavigation();
    }

    private bool IsWithinNavigationPane(DependencyObject source)
    {
        var current = source;
        while (current is not null)
        {
            if (ReferenceEquals(current, NavigationTree) || current is TreeView or TreeViewItem)
                return true;

            current = VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private static bool IsWithinTextInput(DependencyObject source)
    {
        var current = source;
        while (current is not null)
        {
            if (current is TextBox or AutoSuggestBox or PasswordBox or RichEditBox)
                return true;

            current = VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private bool IsWithinFileListArea(DependencyObject source)
    {
        var current = source;
        while (current is not null)
        {
            if (ReferenceEquals(current, FileList))
                return true;

            current = VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private TabContentViewModel? _subscribedTab;

    private void BackButton_Click(object sender, RoutedEventArgs e) => _ = ViewModel.GoBackAsync();

    private async Task OpenFolderFromShortcutFlyoutAsync(string path)
    {
        if (_isShortcutOpenInProgress)
            return;

        _isShortcutOpenInProgress = true;

        CloseShortcutLauncherFlyoutAndFocusRightPane();
        try
        {
            await OpenFolderOrActivateExistingTabAsync(path);
        }
        finally
        {
            _isShortcutOpenInProgress = false;
        }
    }
    private void ForwardButton_Click(object sender, RoutedEventArgs e) => _ = ViewModel.GoForwardAsync();
    private void UpButton_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(ViewModel.ActiveTab?.CurrentPath))
        {
            FileList.PrepareSelectionAfterGoUp(ViewModel.ActiveTab.CurrentPath);
        }

        _ = ViewModel.GoUpAsync();
    }
    private void RefreshButton_Click(object sender, RoutedEventArgs e) => _ = RefreshActiveTabAsync();

    private async Task RefreshActiveTabAsync()
    {
        if (ViewModel.ActiveTab is null || string.IsNullOrWhiteSpace(ViewModel.ActiveTab.CurrentPath))
            return;

        await ViewModel.RefreshAsync();
        UpdateStatusBar();
    }

    private void SubscribeToTabNavigation(TabContentViewModel? tab)
    {
        if (_subscribedTab is not null)
        {
            _subscribedTab.Navigated -= OnActiveTabNavigated;
        }

        _subscribedTab = tab;

        if (tab is not null)
        {
            tab.Navigated += OnActiveTabNavigated;
        }
    }

    private void OnTabPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not TabContentViewModel tab)
            return;

        if (string.Equals(e.PropertyName, nameof(TabContentViewModel.TabTitle), StringComparison.Ordinal))
        {
            _dispatcherQueue.TryEnqueue(() =>
            {
                RefreshTabHeader(tab);

                if (ReferenceEquals(ViewModel.ActiveTab, tab))
                {
                    UpdateWindowTitle(tab);
                }
            });

            return;
        }

        if (string.Equals(e.PropertyName, nameof(TabContentViewModel.CurrentPath), StringComparison.Ordinal))
        {
            _dispatcherQueue.TryEnqueue(() =>
            {
                RefreshTabHeader(tab);

                if (ReferenceEquals(ViewModel.ActiveTab, tab))
                {
                    UpdateAddressBar(tab.CurrentPath);
                    UpdateSearchPlaceholder(tab.CurrentPath);
                }
            });

            return;
        }

        if (string.Equals(e.PropertyName, nameof(TabContentViewModel.IsPinned), StringComparison.Ordinal))
        {
            _dispatcherQueue.TryEnqueue(() => RefreshTabHeader(tab));
        }
    }

    private void OnActiveTabNavigated(object? sender, string path)
    {
        _dispatcherQueue.TryEnqueue(() =>
        {
            if (ViewModel.ActiveTab is not null)
            {
                UpdateWindowTitle(ViewModel.ActiveTab);
                RefreshTabHeader(ViewModel.ActiveTab);
            }

            FileList.SetSearchMode(false);
            UpdateAddressBar(path);
            UpdateStatusBar();
            UpdateSearchPlaceholder(path);
        });
    }

    private void RefreshVisibleTabHeaders()
    {
        foreach (var tab in ViewModel.Tabs)
        {
            RefreshTabHeader(tab);
        }
    }

    private void RefreshTabHeader(TabContentViewModel tab)
    {
        if (!_tabItemsById.TryGetValue(tab.TabState.Id, out var tabViewItem))
            return;

        tabViewItem.DataContext = tab;
        tabViewItem.Header = tab.TabTitle;
        tabViewItem.IsClosable = !tab.IsPinned;
        ToolTipService.SetToolTip(tabViewItem, tab.CurrentPath);
    }

    private void UpdateWindowTitle(TabContentViewModel? tab)
    {
        var tabTitle = tab?.TabTitle;
        if (string.IsNullOrWhiteSpace(tabTitle))
        {
            Title = BaseWindowTitle;
            return;
        }

        Title = $"{tabTitle} - {BaseWindowTitle}";
    }

    private void AddressBar_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        var text = args.ChosenSuggestion as string ?? args.QueryText ?? sender.Text;
        if (!string.IsNullOrWhiteSpace(text))
        {
            _ = ViewModel.NavigateToAddressAsync(text);
        }
        // Switch back to breadcrumb view
        AddressBar.Visibility = Visibility.Collapsed;
        BreadcrumbScroll.Visibility = Visibility.Visible;
    }

    private void AddressBar_SuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
    {
        if (args.SelectedItem is string path)
        {
            sender.Text = path;
        }
    }

    private void AddressBar_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
        {
            var text = sender.Text;
            if (text.Length > 2 && Directory.Exists(Path.GetDirectoryName(text)))
            {
                try
                {
                    var dir = Path.GetDirectoryName(text);
                    var partial = Path.GetFileName(text);
                    if (dir is not null)
                    {
                        var suggestions = Directory.EnumerateDirectories(dir, partial + "*")
                            .Take(10)
                            .Select(d => d)
                            .ToList();
                        sender.ItemsSource = suggestions;
                    }
                }
                catch
                {
                    sender.ItemsSource = null;
                }
            }
        }
    }

    private void AddressBar_LostFocus(object sender, RoutedEventArgs e)
    {
        // Delay slightly so suggestion clicks and QuerySubmitted can fire first
        _dispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            if (AddressBar.FocusState == FocusState.Unfocused)
            {
                AddressBar.Visibility = Visibility.Collapsed;
                BreadcrumbScroll.Visibility = Visibility.Visible;
            }
        });
    }

    private void AddressBarArea_Tapped(object sender, TappedRoutedEventArgs e)
    {
        // If user clicked a breadcrumb segment button, let the button Click handler navigate instead
        var source = e.OriginalSource as DependencyObject;
        while (source is not null && !ReferenceEquals(source, sender))
        {
            if (source is Button) return;
            source = VisualTreeHelper.GetParent(source);
        }
        ShowAddressBarEditMode();
    }

    private void ShowAddressBarEditMode()
    {
        BreadcrumbScroll.Visibility = Visibility.Collapsed;
        AddressBar.Visibility = Visibility.Visible;
        AddressBar.Text = NormalizeDisplayPath(ViewModel.ActiveTab?.CurrentPath ?? "");
        AddressBar.Focus(FocusState.Programmatic);

        // Find the inner TextBox once visible and select all text
        _dispatcherQueue.TryEnqueue(() =>
        {
            if (FindChildOfType<TextBox>(AddressBar) is TextBox tb)
            {
                tb.Focus(FocusState.Programmatic);
                tb.SelectAll();
            }
        });
    }

    private void AddressBar_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Escape)
        {
            AddressBar.Visibility = Visibility.Collapsed;
            BreadcrumbScroll.Visibility = Visibility.Visible;
            e.Handled = true;
        }
    }

    private static T? FindChildOfType<T>(DependencyObject parent) where T : DependencyObject
    {
        var count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T match) return match;
            var found = FindChildOfType<T>(child);
            if (found is not null) return found;
        }
        return null;
    }

    private void UpdateAddressBar(string path)
    {
        path = NormalizeDisplayPath(path);

        // Update hidden edit box text
        AddressBar.Text = path;

        // Build breadcrumb buttons
        BreadcrumbPanel.Children.Clear();

        if (string.IsNullOrEmpty(path)) return;

        // Split the path into segments
        var segments = new List<(string DisplayName, string FullPath)>();
        var current = path.TrimEnd(Path.DirectorySeparatorChar);

        while (!string.IsNullOrEmpty(current))
        {
            var name = Path.GetFileName(current);
            if (string.IsNullOrEmpty(name))
            {
                // Drive root like "C:\"
                name = current;
            }
            segments.Insert(0, (name, current.EndsWith(":") ? current + "\\" : current));

            var parent = Path.GetDirectoryName(current);
            if (parent == current) break; // root
            current = parent!;
        }

        for (int i = 0; i < segments.Count; i++)
        {
            var (displayName, fullPath) = segments[i];

            // Add chevron separator before each segment (except first)
            if (i > 0)
            {
                BreadcrumbPanel.Children.Add(CreateBreadcrumbSeparatorButton(fullPath));
            }

            // Segment button
            var btn = new Button
            {
                Content = displayName,
                Tag = fullPath,
                Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent),
                BorderThickness = new Thickness(0),
                Padding = new Thickness(6, 2, 6, 2),
                MinHeight = 24,
                MinWidth = 0,
                FontSize = 12,
                VerticalAlignment = VerticalAlignment.Center,
                CornerRadius = new CornerRadius(4)
            };
            btn.Click += BreadcrumbSegment_Click;
            BreadcrumbPanel.Children.Add(btn);
        }

        // Keep trailing chevron visible (Folder > Subfolder > Current >)
        BreadcrumbPanel.Children.Add(CreateBreadcrumbSeparatorButton(segments[^1].FullPath));

        // Scroll to rightmost segment
        _dispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            BreadcrumbScroll.ChangeView(BreadcrumbScroll.ScrollableWidth, null, null);
        });
    }

    private void BreadcrumbSegment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string fullPath)
        {
            _ = ViewModel.NavigateToAddressAsync(fullPath);
        }
    }

    private Button CreateBreadcrumbSeparatorButton(string segmentPath)
    {
        var separatorButton = new Button
        {
            Tag = segmentPath,
            Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent),
            BorderThickness = new Thickness(0),
            Padding = new Thickness(2, 0, 2, 0),
            MinHeight = 24,
            MinWidth = 18,
            VerticalAlignment = VerticalAlignment.Center,
            CornerRadius = new CornerRadius(4),
            Content = new FontIcon
            {
                FontFamily = new Microsoft.UI.Xaml.Media.FontFamily("Segoe Fluent Icons"),
                Glyph = "\uE76C",
                FontSize = 10,
                Opacity = 0.6,
                VerticalAlignment = VerticalAlignment.Center
            }
        };

        separatorButton.Click += BreadcrumbSeparator_Click;
        return separatorButton;
    }

    private async void BreadcrumbSeparator_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement anchor || anchor.Tag is not string segmentPath)
            return;

        var requestId = Interlocked.Increment(ref _breadcrumbSiblingRequestId);

        _breadcrumbSiblingCts?.Cancel();
        _breadcrumbSiblingCts?.Dispose();
        _breadcrumbSiblingCts = new CancellationTokenSource();

        var ct = _breadcrumbSiblingCts.Token;
        var flyout = new MenuFlyout();
        flyout.Items.Add(CreateDetailsViewFlyoutItem("Loading siblings...", isEnabled: false));
        flyout.ShowAt(anchor);

        var stopwatch = Stopwatch.StartNew();

        try
        {
            var siblings = await GetSiblingDirectoriesAsync(segmentPath, ct);

            if (requestId != _breadcrumbSiblingRequestId)
                return;

            flyout.Items.Clear();

            if (siblings.Count == 0)
            {
                flyout.Items.Add(CreateDetailsViewFlyoutItem("No sibling folders", isEnabled: false));
                return;
            }

            foreach (var siblingPath in siblings)
            {
                var item = CreateDetailsViewFlyoutItem(GetPathLeafLabel(siblingPath));

                ToolTipService.SetToolTip(item, siblingPath);

                item.Click += (_, _) =>
                {
                    _ = ViewModel.NavigateToAddressAsync(siblingPath);
                };

                flyout.Items.Add(item);
            }
        }
        catch (OperationCanceledException)
        {
            // Superseded by a newer separator click.
        }
        catch (Exception ex)
        {
            flyout.Items.Clear();
            flyout.Items.Add(CreateDetailsViewFlyoutItem($"Failed: {ex.Message}", isEnabled: false));
        }
        finally
        {
            stopwatch.Stop();
            Debug.WriteLine($"Breadcrumb sibling load elapsed: {stopwatch.ElapsedMilliseconds} ms for {segmentPath}");
        }
    }

    private async Task<IReadOnlyList<string>> GetSiblingDirectoriesAsync(string segmentPath, CancellationToken ct)
    {
        var normalizedSegmentPath = NormalizePathForLookup(segmentPath);
        if (_breadcrumbSiblingCache.TryGetValue(normalizedSegmentPath, out var cached))
            return cached;

        var siblings = await Task.Run(() => EnumerateSiblingDirectories(normalizedSegmentPath), ct);
        _breadcrumbSiblingCache[normalizedSegmentPath] = siblings;
        return siblings;
    }

    private static IReadOnlyList<string> EnumerateSiblingDirectories(string segmentPath)
    {
        var root = Path.GetPathRoot(segmentPath);
        if (!string.IsNullOrEmpty(root) && string.Equals(root, segmentPath, StringComparison.OrdinalIgnoreCase))
        {
            return DriveInfo
                .GetDrives()
                .Select(drive => drive.Name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        var trimmed = segmentPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var parent = Path.GetDirectoryName(trimmed);

        if (string.IsNullOrWhiteSpace(parent) || !Directory.Exists(parent))
            return [];

        try
        {
            return Directory
                .EnumerateDirectories(parent)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static string NormalizePathForLookup(string path)
    {
        try
        {
            return TabContentViewModel.NormalizeNavigationPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return path.Trim();
        }
    }

    private static string GetPathLeafLabel(string path)
    {
        var trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var leaf = Path.GetFileName(trimmed);
        return string.IsNullOrWhiteSpace(leaf) ? path : leaf;
    }

    private void UpdateSearchPlaceholder(string path)
    {
        path = NormalizeDisplayPath(path);
        var folderName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar));
        SearchBox.PlaceholderText = string.IsNullOrEmpty(folderName) ? "Search..." : $"Search {folderName}";
    }

    private static string NormalizeDisplayPath(string path)
    {
        if (string.IsNullOrEmpty(path))
            return path;

        if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            return @"\\" + path[8..];

        if (path.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            return path[4..];

        return path;
    }

    private void SearchBox_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        var query = args.QueryText ?? sender.Text;
        if (string.IsNullOrWhiteSpace(query))
        {
            // Empty search — go back to normal directory listing
            FileList.SetSearchMode(false);
            _ = ViewModel.ClearSearchAsync();
        }
        else
        {
            FileList.SetSearchMode(true);
            _ = ViewModel.SearchAsync(query);
        }
    }

    public void ShowShortcutLauncherFlyout(FrameworkElement anchor)
    {
        try
        {
            StartShortcutCachePreload();
            var flyout = GetOrCreateShortcutLauncherFlyout();
            flyout.ShowAt(anchor);
        }
        catch (Exception ex)
        {
            ShowError($"Failed to open shortcuts launcher: {ex.Message}");
        }
    }

    public IReadOnlyList<ShortcutHubRootItem> GetShortcutHubRootItems()
    {
        StartShortcutCachePreload();
        var entries = GetShortcutRootEntriesForFlyout();
        return entries
            .Select(entry => new ShortcutHubRootItem(
                entry.DisplayName,
                entry.FolderPath,
                entry.ShortcutPath,
                !string.IsNullOrWhiteSpace(entry.FolderPath)))
            .ToList();
    }

    public void ShowShortcutHubChildrenFlyout(ShortcutHubRootItem root, FrameworkElement anchor)
    {
        try
        {
            StartShortcutCachePreload();

            if (root.IsFolder && !string.IsNullOrWhiteSpace(root.FolderPath))
            {
                if (!_hubChildrenFlyoutCache.TryGetValue(root.FolderPath, out var flyout))
                {
                    flyout = BuildShortcutHubChildrenFlyout(root.DisplayName, root.FolderPath);
                    _hubChildrenFlyoutCache[root.FolderPath] = flyout;
                }
                flyout.ShowAt(anchor);
                return;
            }

            if (!string.IsNullOrWhiteSpace(root.ShortcutPath))
            {
                OpenShortcutTarget(root.ShortcutPath);
            }
        }
        catch (Exception ex)
        {
            ShowError($"Failed to open favorites entry: {ex.Message}");
        }
    }

    private MenuFlyout GetOrCreateShortcutLauncherFlyout()
    {
        _shortcutLauncherFlyout ??= BuildShortcutLauncherFlyout();
        return _shortcutLauncherFlyout;
    }

    private MenuFlyout BuildShortcutLauncherFlyout()
    {
        var flyout = new MenuFlyout();

        EnsureShortcutsFolderExists();

        if (!Directory.Exists(_shortcutsFolderPath))
        {
            flyout.Items.Add(CreateDetailsViewFlyoutItem("Shortcuts folder not found", isEnabled: false));
            return flyout;
        }

        var entries = GetShortcutRootEntriesForFlyout();
        if (entries.Count == 0)
        {
            flyout.Items.Add(CreateDetailsViewFlyoutItem("No shortcuts found", isEnabled: false));
        }
        else
        {
            foreach (var entry in entries)
            {
                AddShortcutRootEntry(flyout, entry);
            }
        }

        flyout.Items.Add(new MenuFlyoutSeparator());

        var openShortcutsFolder = CreateDetailsViewFlyoutItem("Open shortcuts folder in new tab", "\uE8B7");
        openShortcutsFolder.Click += (_, _) =>
        {
            _ = OpenFolderFromShortcutFlyoutAsync(_shortcutsFolderPath);
        };
        flyout.Items.Add(openShortcutsFolder);

        return flyout;
    }

    private void StartShortcutCachePreload()
    {
        lock (_shortcutCacheLock)
        {
            if (_shortcutPreloadStarted)
                return;

            _shortcutPreloadStarted = true;
        }

        _ = Task.Run(PreloadShortcutCacheAsync);
    }

    private Task PreloadShortcutCacheAsync()
    {
        var totalStopwatch = Stopwatch.StartNew();
        try
        {
            var rootEntries = LoadShortcutRootEntries();

            // Phase 1: quickly cache root folders' immediate children so all first-layer submenus are ready.
            var phase1Stopwatch = Stopwatch.StartNew();
            var shallowEntries = new Dictionary<string, ShortcutFolderCacheData>(StringComparer.OrdinalIgnoreCase);
            foreach (var root in rootEntries)
            {
                if (string.IsNullOrWhiteSpace(root.FolderPath))
                    continue;

                shallowEntries[root.FolderPath] = BuildFolderCacheData(root.FolderPath, includeDescendants: false);
            }

            lock (_shortcutCacheLock)
            {
                _shortcutRootCache = rootEntries;
                foreach (var (path, data) in shallowEntries)
                {
                    _shortcutFolderCache[path] = data;
                }
            }
            phase1Stopwatch.Stop();

            Debug.WriteLine($"Shortcut cache preload phase1 elapsed: {phase1Stopwatch.ElapsedMilliseconds} ms, roots={rootEntries.Count}");

            _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, () =>
            {
                _ = GetOrCreateShortcutLauncherFlyout();
            });

            // Phase 2: deep preload every reachable layer in background so submenu expansion is instant.
            QueueDeepShortcutCachePreload(rootEntries);

            // Phase 3: pre-scan the directory contents of every shortcut target + 2 levels of
            // subdirectories so the first NavigateAsync to any of those paths costs 0 ms disk I/O.
            StartSnapshotPrefetch(rootEntries);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Shortcut cache preload failed: {ex.Message}");
        }
        finally
        {
            totalStopwatch.Stop();
            Debug.WriteLine($"Shortcut cache preload total elapsed: {totalStopwatch.ElapsedMilliseconds} ms");
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Scans every shortcut target folder plus 2 levels of subfolders on a background thread
    /// and stores the results in <see cref="IDirectorySnapshotCache"/>.  The first
    /// NavigateAsync call to any of those paths consumes the snapshot and shows items
    /// instantly with 0 ms disk I/O.
    /// </summary>
    private void StartSnapshotPrefetch(IReadOnlyList<ShortcutRootEntryData> rootEntries)
    {
        if (_snapshotPrefetchStarted)
            return;

        _snapshotPrefetchStarted = true;

        _ = Task.Run(async () =>
        {
            var sw = Stopwatch.StartNew();
            var scanned = 0;

            // Scan level by level (BFS), all paths at a given depth in parallel.
            // EnumerateDirectoryAsync is synchronous under the hood (Win32 FindFirstFileExW).
            // Task.WhenAll(Select(async...)) does NOT give real parallelism for synchronous
            // iterators — all lambdas run on the same thread.  Parallel.ForEachAsync schedules
            // each iteration on a separate thread-pool thread, giving genuine I/O parallelism.
            const int SnapshotMaxDepth = 8;
            const int ScanParallelism = 8; // concurrent Win32 scans per level

            var currentLevel = new List<string>();
            foreach (var root in rootEntries)
            {
                if (!string.IsNullOrWhiteSpace(root.FolderPath))
                    currentLevel.Add(root.FolderPath);
            }

            Debug.WriteLine($"[SnapshotPrefetch] starting  roots={currentLevel.Count}  maxDepth={SnapshotMaxDepth}");

            for (int depth = 0; depth <= SnapshotMaxDepth && currentLevel.Count > 0; depth++)
            {
                var nextLevel = new ConcurrentBag<string>();
                int levelScanned = 0;
                int capturedDepth = depth;

                await Parallel.ForEachAsync(
                    currentLevel,
                    new ParallelOptions { MaxDegreeOfParallelism = ScanParallelism },
                    async (dirPath, ct) =>
                    {
                        try
                        {
                            var entries = new List<FileExplorer.Core.Models.FileSystemEntry>(capacity: 64);
                            await foreach (var entry in _fileSystemService.EnumerateDirectoryAsync(dirPath, ct))
                            {
                                entry.TypeDescription = string.Empty;
                                entries.Add(entry);
                            }

                            _snapshotCache.Store(dirPath, entries);
                            Interlocked.Increment(ref levelScanned);

                            if (capturedDepth < SnapshotMaxDepth)
                            {
                                foreach (var sub in GetSubfolders(dirPath, ShortcutMenuMaxSubfolders))
                                    nextLevel.Add(sub.FullName);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[SnapshotPrefetch] skip {dirPath}: {ex.Message}");
                        }
                    });

                scanned += levelScanned;
                Debug.WriteLine($"[SnapshotPrefetch] depth={capturedDepth} done  scanned={scanned}  queued={nextLevel.Count}  ms={sw.ElapsedMilliseconds}");
                currentLevel = nextLevel.ToList();
            }

            sw.Stop();
            Debug.WriteLine($"[SnapshotPrefetch] done  scanned={scanned}  ms={sw.ElapsedMilliseconds}");
        });
    }

    private void QueueDeepShortcutCachePreload(IReadOnlyList<ShortcutRootEntryData>? rootEntries = null)
    {
        lock (_shortcutCacheLock)
        {
            if (_shortcutDeepPreloadStarted)
                return;

            _shortcutDeepPreloadStarted = true;
        }

        _ = Task.Run(() =>
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                var roots = rootEntries ?? GetShortcutRootEntriesForFlyout();
                var deepEntries = new Dictionary<string, ShortcutFolderCacheData>(StringComparer.OrdinalIgnoreCase);

                foreach (var root in roots)
                {
                    if (string.IsNullOrWhiteSpace(root.FolderPath))
                        continue;

                    PopulateFolderCacheRecursive(root.FolderPath, 0, deepEntries);
                }

                lock (_shortcutCacheLock)
                {
                    foreach (var (path, data) in deepEntries)
                    {
                        _shortcutFolderCache[path] = data;
                    }
                }

                Debug.WriteLine($"Shortcut cache preload phase2 elapsed: {stopwatch.ElapsedMilliseconds} ms, cached={deepEntries.Count}");

                // Invalidate the cached flyout objects so the next open re-builds from the
                // now-complete deep cache.  Do NOT rebuild here on the UI thread — building
                // deeply-nested MenuFlyoutSubItem trees for every level at once would block
                // the UI thread for seconds (one XAML object per cached entry, recursively).
                _dispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
                {
                    _shortcutLauncherFlyout = null;
                    _hubChildrenFlyoutCache.Clear();
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Shortcut cache preload phase2 failed: {ex.Message}");
            }
            finally
            {
                stopwatch.Stop();
            }
        });
    }

    private List<ShortcutRootEntryData> GetShortcutRootEntriesForFlyout()
    {
        lock (_shortcutCacheLock)
        {
            if (_shortcutRootCache.Count > 0)
                return [.. _shortcutRootCache];
        }

        var roots = LoadShortcutRootEntries();
        lock (_shortcutCacheLock)
        {
            _shortcutRootCache = roots;
        }

        return roots;
    }

    private List<ShortcutRootEntryData> LoadShortcutRootEntries()
    {
        var rootEntries = new List<ShortcutRootEntryData>();
        var entries = GetShortcutRootEntries(_shortcutsFolderPath, ShortcutMenuMaxItems);

        foreach (var entry in entries)
        {
            switch (entry)
            {
                case DirectoryInfo directory:
                    rootEntries.Add(new ShortcutRootEntryData(directory.Name, directory.FullName, null));
                    break;

                case FileInfo file when file.Extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase):
                {
                    var displayName = Path.GetFileNameWithoutExtension(file.Name);
                    var targetPath = ResolveShortcutTargetPath(file.FullName);

                    if (!string.IsNullOrWhiteSpace(targetPath) && Directory.Exists(targetPath))
                    {
                        rootEntries.Add(new ShortcutRootEntryData(displayName, targetPath, file.FullName));
                    }
                    else
                    {
                        rootEntries.Add(new ShortcutRootEntryData(displayName, null, file.FullName));
                    }

                    break;
                }
            }
        }

        return rootEntries;
    }

    private static IReadOnlyList<FileSystemInfo> GetShortcutRootEntries(string rootPath, int maxItems)
    {
        try
        {
            var info = new DirectoryInfo(rootPath);
            return info
                .EnumerateFileSystemInfos()
                .Where(item => item is DirectoryInfo ||
                               (item is FileInfo file && file.Extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase)))
                .OrderBy(item => item.Name)
                .Take(maxItems)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private void AddShortcutRootEntry(MenuFlyout flyout, ShortcutRootEntryData entry)
    {
        if (!string.IsNullOrWhiteSpace(entry.FolderPath))
        {
            flyout.Items.Add(CreateFolderSubMenu(entry.DisplayName, entry.FolderPath, 0));
            return;
        }

        if (!string.IsNullOrWhiteSpace(entry.ShortcutPath))
        {
            var item = CreateDetailsViewFlyoutItem(entry.DisplayName);
            ToolTipService.SetToolTip(item, entry.ShortcutPath);
            item.Click += (_, _) => OpenShortcutTarget(entry.ShortcutPath);
            flyout.Items.Add(item);
        }
    }

    private MenuFlyout BuildShortcutHubChildrenFlyout(string displayName, string folderPath)
    {
        var flyout = new MenuFlyout();

        var openCurrentFolder = CreateDetailsViewFlyoutItem($"Open {displayName}", "\uE8B7");
        openCurrentFolder.Click += (_, _) =>
        {
            _ = OpenFolderFromShortcutFlyoutAsync(folderPath);
        };
        flyout.Items.Add(openCurrentFolder);
        flyout.Items.Add(new MenuFlyoutSeparator());

        ShortcutFolderCacheData folderData;
        if (!TryGetShortcutFolderCache(folderPath, out folderData))
        {
            folderData = BuildFolderCacheData(folderPath, includeDescendants: false);
            lock (_shortcutCacheLock)
            {
                _shortcutFolderCache[folderPath] = folderData;
            }
        }

        // If the cache entry exists but is empty (transient warm-up race where phase-1 ran
        // but found 0 items before the filesystem returned results), do a targeted sync refresh
        // so the flyout is never shown blank on first open.
        if (!HasFolderCacheItems(folderData) && !folderData.IsFullyLoaded)
        {
            var refreshed = BuildFolderCacheData(folderPath, includeDescendants: false);
            folderData = refreshed;
            lock (_shortcutCacheLock)
            {
                _shortcutFolderCache[folderPath] = folderData;
            }
        }

        AddFolderItemsToFlyout(flyout, folderData, 1);
        StartFolderDescendantPrefetch(folderPath, 1);

        if (flyout.Items.Count <= 2)
        {
            flyout.Items.Add(CreateDetailsViewFlyoutItem("No items", isEnabled: false));
        }

        return flyout;
    }

    private MenuFlyoutSubItem CreateFolderSubMenu(string displayName, string folderPath, int depth)
    {
        var folderItem = new MenuFlyoutSubItem { Text = displayName };
        ApplyDetailsViewFlyoutSubItemStyle(folderItem);
        ToolTipService.SetToolTip(folderItem, folderPath);
        var context = new ShortcutFolderMenuContext(folderPath, depth);
        folderItem.Tag = context;
        folderItem.AddHandler(UIElement.PointerPressedEvent, new PointerEventHandler(OnShortcutFolderItemPointerPressed), true);

        if (depth >= ShortcutMenuMaxDepth)
            return folderItem;

        // Pre-populate the three outermost levels eagerly from cache (depth 0, 1, 2).
        // This covers:
        //   - Main flyout: root subfolders (depth 0), their children (depth 1), grandchildren (depth 2)
        //   - Hub flyout : hub-root's children (depth 1), their children (depth 2), grandchildren (depth 3)
        // Depth 3+ items use InitializeFolderLazyLoading which reads from the in-memory cache
        // synchronously on PointerEntered — effectively instant once phase-2 preload completes.
        //
        // IMPORTANT: do not raise this limit beyond depth < 3 without careful analysis.
        // Worst case: ShortcutMenuMaxSubfolders^3 = 50^3 = 125 000 XAML objects in one dispatch.
        if (depth < 3 && TryGetShortcutFolderCache(folderPath, out var cachedData))
        {
            if (HasFolderCacheItems(cachedData))
            {
                AddFolderSubMenuItems(folderItem, cachedData, depth + 1);
                context.IsLoaded = true;
                if (!cachedData.IsFullyLoaded)
                    StartFolderDescendantPrefetch(folderPath, depth);
                return folderItem;
            }

            if (cachedData.IsFullyLoaded)
            {
                // Genuinely empty folder.
                context.IsLoaded = true;
                return folderItem;
            }

            // Empty but not fully loaded → warm-up race; fall through to lazy loading.
        }

        InitializeFolderLazyLoading(folderItem);

        return folderItem;
    }

    private void InitializeFolderLazyLoading(MenuFlyoutSubItem folderItem)
    {
        folderItem.Items.Clear();
        folderItem.Items.Add(CreateDetailsViewFlyoutItem("Loading...", isEnabled: false));

        // Use AddHandler with handledEventsToo=true for ALL triggers.
        // WinUI's internal flyout logic marks pointer/focus events as Handled at deeper nesting
        // levels, which silently drops += subscriptions.  AddHandler(..., true) fires regardless.
        // PointerMoved is added as an extra-safe fallback: it fires on any cursor movement over
        // the item row and is harder for WinUI to swallow than PointerEntered.
        PointerEventHandler lazyLoadPointerHandler = (_, _) => _ = EnsureFolderSubMenuLoadedAsync(folderItem);
        folderItem.AddHandler(UIElement.PointerEnteredEvent, lazyLoadPointerHandler, handledEventsToo: true);
        folderItem.AddHandler(UIElement.PointerMovedEvent, lazyLoadPointerHandler, handledEventsToo: true);
        folderItem.GotFocus += (_, _) => _ = EnsureFolderSubMenuLoadedAsync(folderItem);
        folderItem.KeyDown += (_, _) => _ = EnsureFolderSubMenuLoadedAsync(folderItem);
    }

    private void OnShortcutFolderItemPointerPressed(object sender, PointerRoutedEventArgs args)
    {
        if (sender is not MenuFlyoutSubItem folderItem ||
            folderItem.Tag is not ShortcutFolderMenuContext context)
        {
            return;
        }

        var point = args.GetCurrentPoint(folderItem);
        if (!point.Properties.IsLeftButtonPressed)
            return;

        // Only treat clicks on this row; clicking child rows should not open the parent folder.
        var owner = FindOwningMenuFlyoutItem(args.OriginalSource as DependencyObject);
        if (!ReferenceEquals(owner, folderItem))
            return;

        _ = OpenFolderFromShortcutFlyoutAsync(context.FolderPath);

        args.Handled = true;
    }

    private static MenuFlyoutItemBase? FindOwningMenuFlyoutItem(DependencyObject? source)
    {
        var current = source;
        while (current is not null)
        {
            if (current is MenuFlyoutItemBase item)
                return item;

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private async Task EnsureFolderSubMenuLoadedAsync(MenuFlyoutSubItem folderItem)
    {
        if (folderItem.Tag is not ShortcutFolderMenuContext context || context.IsLoaded || context.IsLoading)
            return;

        var stopwatch = Stopwatch.StartNew();
        var source = "unknown";

        context.IsLoading = true;

        try
        {
            if (context.Depth >= ShortcutMenuMaxDepth)
            {
                context.IsLoaded = true;
                folderItem.Items.Clear();
                folderItem.Items.Add(CreateDetailsViewFlyoutItem("No deeper levels", isEnabled: false));
                source = "max-depth";
                return;
            }

            if (TryGetShortcutFolderCache(context.FolderPath, out var cachedData))
            {
                var hasCachedItems = HasFolderCacheItems(cachedData);
                if (hasCachedItems || cachedData.IsFullyLoaded)
                {
                    source = cachedData.IsFullyLoaded ? "cache:full" : "cache:shallow";
                    folderItem.Items.Clear();
                    AddFolderSubMenuItems(folderItem, cachedData, context.Depth + 1);
                    if (folderItem.Items.Count == 0)
                        folderItem.Items.Add(CreateDetailsViewFlyoutItem("No items", isEnabled: false));

                    if (!cachedData.IsFullyLoaded)
                        StartFolderDescendantPrefetch(context.FolderPath, context.Depth + 1);

                    context.IsLoaded = true;
                    return;
                }

                source = "cache:empty-warmup";
            }

            folderItem.Items.Clear();
            folderItem.Items.Add(CreateDetailsViewFlyoutItem("Loading...", isEnabled: false));

            var fallbackData = await Task.Run(() => BuildFolderCacheData(context.FolderPath, includeDescendants: false));
            source = "filesystem:fallback";

            lock (_shortcutCacheLock)
            {
                _shortcutFolderCache[context.FolderPath] = fallbackData;
            }

            StartFolderDescendantPrefetch(context.FolderPath, context.Depth + 1);

            folderItem.Items.Clear();
            AddFolderSubMenuItems(folderItem, fallbackData, context.Depth + 1);
            if (folderItem.Items.Count == 0)
                folderItem.Items.Add(CreateDetailsViewFlyoutItem("No items", isEnabled: false));

            context.IsLoaded = true;
        }
        catch (Exception ex)
        {
            context.IsLoaded = false;
            folderItem.Items.Clear();
            folderItem.Items.Add(CreateDetailsViewFlyoutItem("Failed to load (hover to retry)", isEnabled: false));
            source = "error";
            Debug.WriteLine($"Shortcut submenu load failed: depth={context.Depth}, path={context.FolderPath}, error={ex.Message}");
        }
        finally
        {
            context.IsLoading = false;
            stopwatch.Stop();
            Debug.WriteLine($"Shortcut submenu load elapsed: {stopwatch.ElapsedMilliseconds} ms, depth={context.Depth}, source={source}, path={context.FolderPath}");
        }
    }

    private void AddFolderSubMenuItems(MenuFlyoutSubItem folderItem, ShortcutFolderCacheData data, int childDepth)
    {
        foreach (var subfolder in data.Subfolders)
        {
            folderItem.Items.Add(CreateFolderSubMenu(subfolder.DisplayName, subfolder.FolderPath, childDepth));
        }

        if (data.Subfolders.Count > 0 && data.Files.Count > 0)
            folderItem.Items.Add(new MenuFlyoutSeparator());

        foreach (var file in data.Files)
        {
            var fileItem = CreateDetailsViewFlyoutItem(file.Name);
            ToolTipService.SetToolTip(fileItem, file.FullPath);
            fileItem.Click += (_, _) => OpenFileOrShortcut(file.FullPath);
            folderItem.Items.Add(fileItem);
        }
    }

    private void AddFolderItemsToFlyout(MenuFlyout flyout, ShortcutFolderCacheData data, int childDepth)
    {
        foreach (var subfolder in data.Subfolders)
        {
            flyout.Items.Add(CreateFolderSubMenu(subfolder.DisplayName, subfolder.FolderPath, childDepth));
        }

        if (data.Subfolders.Count > 0 && data.Files.Count > 0)
            flyout.Items.Add(new MenuFlyoutSeparator());

        foreach (var file in data.Files)
        {
            var fileItem = CreateDetailsViewFlyoutItem(file.Name);
            ToolTipService.SetToolTip(fileItem, file.FullPath);
            fileItem.Click += (_, _) => OpenFileOrShortcut(file.FullPath);
            flyout.Items.Add(fileItem);
        }
    }

    private bool TryGetShortcutFolderCache(string folderPath, out ShortcutFolderCacheData data)
    {
        lock (_shortcutCacheLock)
        {
            return _shortcutFolderCache.TryGetValue(folderPath, out data!);
        }
    }

    private static bool HasFolderCacheItems(ShortcutFolderCacheData data)
    {
        return data.Subfolders.Count > 0 || data.Files.Count > 0;
    }

    private ShortcutFolderCacheData BuildFolderCacheData(string folderPath, bool includeDescendants)
    {
        var subfolders = GetSubfolders(folderPath, ShortcutMenuMaxSubfolders);
        var files = GetFiles(folderPath, ShortcutMenuMaxSubfolders);

        return new ShortcutFolderCacheData(
            [.. subfolders.Select(sub => new ShortcutSubfolderData(sub.Name, sub.FullName))],
            [.. files.Select(file => new ShortcutFileData(file.Name, file.FullName))],
            IsFullyLoaded: includeDescendants);
    }

    private void PopulateFolderCacheRecursive(string folderPath, int depth, Dictionary<string, ShortcutFolderCacheData> cache)
    {
        if (depth >= ShortcutMenuMaxDepth)
            return;

        var data = BuildFolderCacheData(folderPath, includeDescendants: true);
        cache[folderPath] = data;

        foreach (var subfolder in data.Subfolders)
        {
            PopulateFolderCacheRecursive(subfolder.FolderPath, depth + 1, cache);
        }
    }

    private void StartFolderDescendantPrefetch(string folderPath, int depth)
    {
        if (depth >= ShortcutMenuMaxDepth)
            return;

        lock (_shortcutCacheLock)
        {
            if (_shortcutPrefetchInProgress.Contains(folderPath))
                return;

            _shortcutPrefetchInProgress.Add(folderPath);
        }

        _ = Task.Run(() =>
        {
            try
            {
                var deepEntries = new Dictionary<string, ShortcutFolderCacheData>(StringComparer.OrdinalIgnoreCase);
                PopulateFolderCacheRecursive(folderPath, depth, deepEntries);

                lock (_shortcutCacheLock)
                {
                    foreach (var (path, data) in deepEntries)
                    {
                        _shortcutFolderCache[path] = data;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Shortcut descendant prefetch failed: {ex.Message}");
            }
            finally
            {
                lock (_shortcutCacheLock)
                {
                    _shortcutPrefetchInProgress.Remove(folderPath);
                }
            }
        });
    }

    private static List<DirectoryInfo> GetSubfolders(string folderPath, int maxCount)
    {
        try
        {
            if (!Directory.Exists(folderPath))
                return [];

            return new DirectoryInfo(folderPath)
                .EnumerateDirectories()
                .OrderBy(d => d.Name)
                .Take(maxCount)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static List<FileInfo> GetFiles(string folderPath, int maxCount)
    {
        try
        {
            if (!Directory.Exists(folderPath))
                return [];

            return new DirectoryInfo(folderPath)
                .EnumerateFiles()
                .OrderBy(f => f.Name)
                .Take(maxCount)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private string ResolveShortcutTargetPath(string shortcutPath)
    {
        try
        {
            var shellService = App.Services.GetRequiredService<IShellIntegrationService>();
            return shellService.ResolveShortcutTarget(shortcutPath) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void OpenShortcutTarget(string shortcutPath)
    {
        var targetPath = ResolveShortcutTargetPath(shortcutPath);
        if (string.IsNullOrWhiteSpace(targetPath))
        {
            OpenFileOrShortcut(shortcutPath);
            return;
        }

        OpenFileOrShortcut(targetPath);
    }

    private void OpenFileOrShortcut(string path)
    {
        try
        {
            if (TryResolveShortcutFolderTarget(path, out var shortcutFolderPath))
            {
                _ = OpenFolderFromShortcutFlyoutAsync(shortcutFolderPath);
                return;
            }

            if (Directory.Exists(path))
            {
                _ = OpenFolderFromShortcutFlyoutAsync(path);
                return;
            }

            var shellService = App.Services.GetRequiredService<IShellIntegrationService>();
            shellService.OpenFile(path);
            CloseShortcutLauncherFlyoutAndFocusRightPane();
        }
        catch (Exception ex)
        {
            ShowError($"Failed to open item: {ex.Message}");
        }
    }

    private bool TryResolveShortcutFolderTarget(string path, out string folderPath)
    {
        folderPath = string.Empty;

        if (!string.Equals(Path.GetExtension(path), ".lnk", StringComparison.OrdinalIgnoreCase))
            return false;

        var resolvedTarget = ResolveShortcutTargetPath(path);
        if (string.IsNullOrWhiteSpace(resolvedTarget))
            return false;

        try
        {
            resolvedTarget = TabContentViewModel.NormalizeNavigationPath(resolvedTarget);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }

        if (!Directory.Exists(resolvedTarget))
            return false;

        folderPath = resolvedTarget;
        return true;
    }

    private void EnsureShortcutsFolderExists()
    {
        try
        {
            if (!Directory.Exists(_shortcutsFolderPath))
            {
                Directory.CreateDirectory(_shortcutsFolderPath);
            }
        }
        catch
        {
            // Best-effort only; UI will display an error item if the path remains unavailable.
        }
    }

    private static string ResolveShortcutsFolderPath()
    {
        var configuredPath = Environment.GetEnvironmentVariable("FILEEXPLORER_SHORTCUTS_FOLDER");
        if (!string.IsNullOrWhiteSpace(configuredPath))
        {
            var expandedConfiguredPath = Environment.ExpandEnvironmentVariables(configuredPath);
            if (Directory.Exists(expandedConfiguredPath))
                return expandedConfiguredPath;
        }

        var userProfileShortcuts = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Shortcuts_to_folders");

        if (Directory.Exists(userProfileShortcuts))
            return userProfileShortcuts;

        var nearbyShortcuts = FindNearbyShortcutsFolder();
        return nearbyShortcuts ?? userProfileShortcuts;
    }

    private static string? FindNearbyShortcutsFolder()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        for (int i = 0; i < 10 && directory is not null; i++)
        {
            var singular = Path.Combine(directory.FullName, "Shortcut_to_folders");
            if (Directory.Exists(singular))
                return singular;

            var plural = Path.Combine(directory.FullName, "Shortcuts_to_folders");
            if (Directory.Exists(plural))
                return plural;

            directory = directory.Parent;
        }

        return null;
    }

    #endregion

    #region Tree View

    private void PopulateTreeView()
    {
        NavigationTree.RootNodes.Clear();
        foreach (var rootVm in ViewModel.TreeView.RootNodes)
        {
            NavigationTree.RootNodes.Add(CreateTreeViewNode(rootVm));
        }
    }

    private TreeViewNode CreateTreeViewNode(TreeNodeViewModel viewModel)
    {
        var node = new TreeViewNode
        {
            Content = viewModel,
            HasUnrealizedChildren = viewModel.HasChildren && !viewModel.IsLoaded,
            // Do NOT set IsExpanded here — WinUI ignores it before the node is in the tree.
        };

        // If children are already loaded, add them now
        if (viewModel.IsLoaded)
        {
            foreach (var child in viewModel.Children)
            {
                // Skip "Loading..." placeholder nodes
                if (string.IsNullOrEmpty(child.Model.Path)) continue;
                node.Children.Add(CreateTreeViewNode(child));
            }
        }

        return node;
    }

    /// <summary>Walk all TreeViewNodes and set IsExpanded from the backing ViewModel AFTER the nodes are in the tree.</summary>
    private static void ApplyExpansionState(IList<TreeViewNode> nodes)
    {
        foreach (var tvNode in nodes)
        {
            if (tvNode.Content is TreeNodeViewModel vm && vm.IsExpanded)
            {
                tvNode.IsExpanded = true;
            }
            if (tvNode.Children.Count > 0)
            {
                ApplyExpansionState(tvNode.Children);
            }
        }
    }

    private void NavigationTree_ItemInvoked(TreeView sender, TreeViewItemInvokedEventArgs args)
    {
        if (args.InvokedItem is TreeViewNode tvn && tvn.Content is TreeNodeViewModel node)
        {
            ViewModel.TreeView.SelectNodeCommand.Execute(node);
        }
    }

    private void NavigationTree_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        try
        {
            if (FindAncestorOfType<TreeViewItem>(e.OriginalSource as DependencyObject) is not TreeViewItem treeViewItem)
                return;

            if (NavigationTree.NodeFromContainer(treeViewItem) is not TreeViewNode tvNode)
                return;

            if (tvNode.Content is not TreeNodeViewModel node)
                return;

            NavigationTree.SelectedNode = tvNode;
            var resolvedPath = ResolveTreeNodePath(node);
            if (string.IsNullOrWhiteSpace(resolvedPath))
                return;

            var flyout = new MenuFlyout();
            AddFolderOpenItems(flyout, resolvedPath);
            if (flyout.Items.Count == 0)
                return;

            flyout.ShowAt(treeViewItem, e.GetPosition(treeViewItem));
            e.Handled = true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Navigation tree right-click menu failed: {ex}");
        }
    }

    private async void NavigationTree_Expanding(TreeView sender, TreeViewExpandingEventArgs args)
    {
        var tvNode = args.Node;
        if (tvNode.Content is TreeNodeViewModel node)
        {
            node.IsExpanded = true;

            if (!node.IsLoaded)
            {
                await ViewModel.TreeView.ExpandNodeAsync(node);

                tvNode.Children.Clear();
                foreach (var child in node.Children)
                {
                    if (string.IsNullOrEmpty(child.Model.Path)) continue;
                    tvNode.Children.Add(CreateTreeViewNode(child));
                }
                tvNode.HasUnrealizedChildren = false;
            }
        }
    }

    private void NavigationTree_Collapsed(TreeView sender, TreeViewCollapsedEventArgs args)
    {
        if (args.Node.Content is TreeNodeViewModel node)
        {
            node.IsExpanded = false;
        }
    }

    private bool _navigationTreeLoaded;
    private string? _pendingTreeSyncPath;

    private void TreeSplitter_ManipulationDelta(object sender, ManipulationDeltaRoutedEventArgs e)
    {
        var newWidth = PreviewColumn.Width.Value - e.Delta.Translation.X;
        if (newWidth >= 220 && newWidth <= 700)
        {
            PreviewColumn.Width = new GridLength(newWidth);
            ViewModel.PreviewPaneWidth = newWidth;
        }
    }

    private void TreeSplitter_PointerEntered(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Grid grid)
        {
            grid.Background = new SolidColorBrush(Microsoft.UI.Colors.LightGray) { Opacity = 0.25 };
            TreeSplitterHint.Visibility = Visibility.Visible;
            SetCursor(LoadCursor(nint.Zero, IDC_SIZEWE));
        }
    }

    private void TreeSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        SetCursor(LoadCursor(nint.Zero, IDC_SIZEWE));
    }

    private void TreeSplitter_PointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Grid grid)
        {
            grid.Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent);
            TreeSplitterHint.Visibility = Visibility.Collapsed;
            SetCursor(LoadCursor(nint.Zero, IDC_ARROW));
        }
    }

    private int _treeSyncRequestId;

    private async Task SyncNavigationTreeAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        var normalizedPath = NormalizeTreeSyncPath(path);

        if (!string.IsNullOrWhiteSpace(normalizedPath) &&
            string.Equals(normalizedPath, _lastSyncedTreePath, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (!_navigationTreeLoaded)
        {
            _pendingTreeSyncPath = path;
            return;
        }

        var requestId = Interlocked.Increment(ref _treeSyncRequestId);
        TreeNodeViewModel? selectedVm;

        try
        {
            selectedVm = await ViewModel.TreeView.SyncWithPathAsync(path);
            _lastSyncedTreePath = normalizedPath;
        }
        catch (Exception ex)
        {
            ShowError($"Tree sync failed: {ex.Message}");
            return;
        }

        _dispatcherQueue.TryEnqueue(() =>
        {
            if (requestId != _treeSyncRequestId)
                return;

            PopulateTreeView();
            ApplyTreeSelection(selectedVm);
        });
    }

    private void ApplyTreeSelection(TreeNodeViewModel? selectedVm)
    {
        ApplyExpansionState(NavigationTree.RootNodes);

        if (selectedVm is null)
        {
            NavigationTree.SelectedNode = null;
            return;
        }

        var tvNode = FindTreeViewNode(NavigationTree.RootNodes, selectedVm);
        if (tvNode is null)
        {
            NavigationTree.SelectedNode = null;
            return;
        }

        ExpandNodePath(tvNode);
        NavigationTree.SelectedNode = tvNode;
    }

    private static string NormalizeTreeSyncPath(string path)
    {
        try
        {
            return TabContentViewModel.NormalizeNavigationPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return path.Trim();
        }
    }

    private static TreeViewNode? FindTreeViewNode(IList<TreeViewNode> nodes, TreeNodeViewModel target)
    {
        foreach (var node in nodes)
        {
            if (node.Content is TreeNodeViewModel vm && vm == target)
                return node;

            var found = FindTreeViewNode(node.Children, target);
            if (found is not null) return found;
        }
        return null;
    }

    public async Task OpenFolderInNewTabAsync(string path)
    {
        await ViewModel.OpenFolderInNewTabAsync(path);
        UpdateTabBindings();

        if (ViewModel.ActiveTab is not null)
        {
            ActivateTab(ViewModel.ActiveTab, syncTabViewSelection: true);
        }

        ActivateRightPaneForKeyboardNavigation();
    }

    public async Task OpenFolderOrActivateExistingTabAsync(string path)
    {
        await ViewModel.OpenFolderOrActivateExistingTabAsync(path);
        UpdateTabBindings();

        if (ViewModel.ActiveTab is not null)
        {
            ActivateTab(ViewModel.ActiveTab, syncTabViewSelection: true);
        }

        ActivateRightPaneForKeyboardNavigation();
    }

    public async Task AddFavoritesAsync(IReadOnlyList<string> paths)
    {
        if (paths.Count == 0)
            return;

        var normalizedPaths = paths
            .Select(NormalizeFavoritePath)
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (normalizedPaths.Count == 0)
            return;

        var favorites = _settingsService.Settings.Favorites;
        var existing = favorites
            .Select(favorite => NormalizeFavoritePath(favorite.Path))
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var changed = false;

        foreach (var favoritePath in normalizedPaths)
        {
            if (existing.Add(favoritePath))
            {
                favorites.Add(new FavoriteEntrySettings { Path = favoritePath });
                changed = true;
            }
        }

        if (!changed)
            return;

        _settingsService.Settings.Favorites = NormalizeAndSortFavorites(favorites);

        await _settingsService.SaveAsync();
        InspectorSidebar.RefreshFavorites();
    }

    public async Task RemoveFavoritesAsync(IReadOnlyList<string> paths)
    {
        if (paths.Count == 0)
            return;

        var normalizedToRemove = paths
            .Select(NormalizeFavoritePath)
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (normalizedToRemove.Count == 0)
            return;

        var favorites = _settingsService.Settings.Favorites;
        var updated = favorites
            .Where(favorite => !normalizedToRemove.Contains(NormalizeFavoritePath(favorite.Path)))
            .ToList();

        if (updated.Count == favorites.Count)
            return;

        _settingsService.Settings.Favorites = NormalizeAndSortFavorites(updated);
        await _settingsService.SaveAsync();
        InspectorSidebar.RefreshFavorites();
    }

    public async Task RenameFavoriteAsync(string path, string? alias)
    {
        var normalizedPath = NormalizeFavoritePath(path);
        if (string.IsNullOrWhiteSpace(normalizedPath))
            return;

        var favorites = _settingsService.Settings.Favorites;
        var favorite = favorites.FirstOrDefault(item =>
            string.Equals(NormalizeFavoritePath(item.Path), normalizedPath, StringComparison.OrdinalIgnoreCase));

        if (favorite is null)
            return;

        favorite.Alias = NormalizeFavoriteAlias(alias, normalizedPath);
        _settingsService.Settings.Favorites = NormalizeAndSortFavorites(favorites);

        await _settingsService.SaveAsync();
        InspectorSidebar.RefreshFavorites();
    }

    public async Task OpenFavoritePathAsync(string path)
    {
        try
        {
            var normalizedPath = NormalizeFavoritePath(path);
            if (string.IsNullOrWhiteSpace(normalizedPath))
                return;

            if (TryResolveShortcutFolderTarget(normalizedPath, out var shortcutFolderPath))
            {
                await OpenFolderOrActivateExistingTabAsync(shortcutFolderPath);
                return;
            }

            if (Directory.Exists(normalizedPath))
            {
                await OpenFolderOrActivateExistingTabAsync(normalizedPath);
                return;
            }

            if (File.Exists(normalizedPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = normalizedPath,
                    UseShellExecute = true,
                    WorkingDirectory = Path.GetDirectoryName(normalizedPath) ?? Environment.CurrentDirectory
                });
                return;
            }

            ShowError($"Favorite path not found: {normalizedPath}");
        }
        catch (Exception ex)
        {
            ShowError($"Failed to open favorite: {ex.Message}");
        }
    }

    public void OpenFolderInNewWindow(string path)
    {
        path = TabContentViewModel.NormalizeNavigationPath(path);
        if (!Directory.Exists(path))
            return;

        var executablePath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(executablePath))
            return;

        Process.Start(new ProcessStartInfo
        {
            FileName = executablePath,
            Arguments = $"\"{path}\"",
            UseShellExecute = true,
            WorkingDirectory = Path.GetDirectoryName(executablePath) ?? Environment.CurrentDirectory
        });
    }

    private void AddFolderOpenItems(MenuFlyout flyout, string path)
    {
        flyout.Items.Add(CreateDetailsViewFlyoutItem("Open", "\uE8E5"));
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            _ = ViewModel.NavigateToAddressAsync(path);
        };

        flyout.Items.Add(CreateDetailsViewFlyoutItem("Open in new tab", "\uE7C3"));
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            _ = OpenFolderInNewTabAsync(path);
        };

        flyout.Items.Add(CreateDetailsViewFlyoutItem("Open in new window", "\uE8A7"));
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            OpenFolderInNewWindow(path);
        };
    }

    private static string NormalizeFavoritePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        var candidate = path.Trim().Trim('"');

        try
        {
            candidate = TabContentViewModel.NormalizeNavigationPath(candidate);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            // Keep best-effort raw path for display/open attempts.
        }

        return candidate;
    }

    private static List<FavoriteEntrySettings> NormalizeAndSortFavorites(IEnumerable<FavoriteEntrySettings> favorites)
    {
        var deduped = new Dictionary<string, FavoriteEntrySettings>(StringComparer.OrdinalIgnoreCase);
        foreach (var favorite in favorites)
        {
            var normalizedPath = NormalizeFavoritePath(favorite.Path);
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

        return deduped.Values
            .OrderBy(GetFavoriteSortLabel, StringComparer.OrdinalIgnoreCase)
            .ThenBy(favorite => favorite.Path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string GetFavoriteSortLabel(FavoriteEntrySettings favorite)
    {
        return string.IsNullOrWhiteSpace(favorite.Alias)
            ? GetPathLeafLabel(favorite.Path)
            : favorite.Alias;
    }

    private static string? NormalizeFavoriteAlias(string? alias, string favoritePath)
    {
        if (string.IsNullOrWhiteSpace(alias))
            return null;

        var trimmed = alias.Trim();
        return string.Equals(trimmed, GetPathLeafLabel(favoritePath), StringComparison.OrdinalIgnoreCase)
            ? null
            : trimmed;
    }

    private static TreeViewItem? FindAncestorOfType<TreeViewItem>(DependencyObject? element) where TreeViewItem : DependencyObject
    {
        while (element is not null)
        {
            if (element is TreeViewItem match)
                return match;

            element = VisualTreeHelper.GetParent(element);
        }

        return null;
    }

    private static string? ResolveTreeNodePath(TreeNodeViewModel node)
    {
        var path = node.Model.Path;
        if (string.IsNullOrWhiteSpace(path))
            return null;

        if (!path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
            return path;

        var shellService = App.Services.GetRequiredService<IShellIntegrationService>();
        return shellService.ResolveShellPath(path);
    }

    private void ExpandNodePath(TreeViewNode node)
    {
        var path = new Stack<TreeViewNode>();
        for (TreeViewNode? current = node; current is not null; current = current.Parent)
        {
            path.Push(current);
        }

        while (path.Count > 0)
        {
            var current = path.Pop();
            current.IsExpanded = true;
            NavigationTree.Expand(current);
        }
    }

    private void NavigationTree_Loaded(object sender, RoutedEventArgs e)
    {
        _navigationTreeLoaded = true;

        ApplyExpansionState(NavigationTree.RootNodes);

        if (_pendingSelectedTreeNode is not null)
        {
            ApplyTreeSelection(_pendingSelectedTreeNode);
            _pendingSelectedTreeNode = null;
            return;
        }

        if (string.IsNullOrWhiteSpace(_pendingTreeSyncPath))
            return;

        var path = _pendingTreeSyncPath;
        _pendingTreeSyncPath = null;
        _ = SyncNavigationTreeAsync(path);
    }

    #endregion

    #region Command Bar Handlers

    private void NewFolder_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.ActiveTab is not null)
            _ = ViewModel.ActiveTab.CreateNewFolderAsync();
    }

    private void Cut_Click(object sender, RoutedEventArgs e) => CutCopyToClipboard(isMove: true);
    private void Copy_Click(object sender, RoutedEventArgs e) => CutCopyToClipboard(isMove: false);

    public void CutToClipboard() => CutCopyToClipboard(isMove: true);
    public void CopyToClipboard() => CutCopyToClipboard(isMove: false);
    private void Paste_Click(object sender, RoutedEventArgs e) => _ = PasteFromClipboardAsync();
    private void Rename_Click(object sender, RoutedEventArgs e) => ViewModel.ActiveTab?.StartRename();

    private void Share_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.ActiveTab?.SelectedItems.Count > 0)
        {
            var path = ViewModel.ActiveTab.SelectedItems[0].Entry.FullPath;
            App.Services.GetRequiredService<Core.Interfaces.IShellIntegrationService>().ShowShareDialog(Hwnd, path);
        }
    }

    private void Delete_Click(object sender, RoutedEventArgs e) => _ = DeleteSelectedWithConfirmationAsync();
    private void SelectAll_Click(object sender, RoutedEventArgs e) => ViewModel.ActiveTab?.SelectAll();
    private void Undo_Click(object sender, RoutedEventArgs e) => _ = ViewModel.UndoAsync();
    private void Properties_Click(object sender, RoutedEventArgs e) => ViewModel.ShowProperties(Hwnd);

    private void SortByName_Click(object sender, RoutedEventArgs e) => ViewModel.ActiveTab?.SetSortCommand.Execute("Name");
    private void SortByDate_Click(object sender, RoutedEventArgs e) => ViewModel.ActiveTab?.SetSortCommand.Execute("DateModified");
    private void SortByType_Click(object sender, RoutedEventArgs e) => ViewModel.ActiveTab?.SetSortCommand.Execute("Type");
    private void SortBySize_Click(object sender, RoutedEventArgs e) => ViewModel.ActiveTab?.SetSortCommand.Execute("Size");

    private void TogglePreviewPane_Click(object sender, RoutedEventArgs e)
    {
        // Preview pane has been replaced by the inspector sidebar in this mode.
        PreviewColumn.Width = new GridLength(ViewModel.PreviewPaneWidth > 0 ? ViewModel.PreviewPaneWidth : 320);
    }

    private void ToggleHiddenFiles_Click(object sender, RoutedEventArgs e) => ViewModel.ToggleHiddenFiles();
    private void ToggleFileExtensions_Click(object sender, RoutedEventArgs e) => ViewModel.ToggleFileExtensions();

    #endregion

    #region Clipboard

    private void CutCopyToClipboard(bool isMove)
    {
        if (ViewModel.ActiveTab is null) return;
        var selectedPaths = ViewModel.ActiveTab.SelectedItems.Select(i => i.Entry.FullPath).ToList();
        if (selectedPaths.Count == 0) return;

        var package = new DataPackage();
        package.RequestedOperation = isMove ? DataPackageOperation.Move : DataPackageOperation.Copy;

        var files = new List<IStorageItem>();
        foreach (var path in selectedPaths)
        {
            try
            {
                if (Directory.Exists(path))
                    files.Add(StorageFolder.GetFolderFromPathAsync(path).AsTask().Result);
                else if (File.Exists(path))
                    files.Add(StorageFile.GetFileFromPathAsync(path).AsTask().Result);
            }
            catch { /* Skip inaccessible items */ }
        }

        if (files.Count > 0)
        {
            package.SetStorageItems(files);
            Clipboard.SetContent(package);
            ViewModel.HasClipboardContent = true;
        }
    }

    private async Task PasteFromClipboardAsync()
    {
        if (ViewModel.ActiveTab is null) return;
        var destPath = ViewModel.ActiveTab.CurrentPath;
        if (string.IsNullOrEmpty(destPath)) return;

        try
        {
            var content = Clipboard.GetContent();
            if (content.Contains(StandardDataFormats.StorageItems))
            {
                var items = await content.GetStorageItemsAsync();
                var sourcePaths = items.Select(i => i.Path).ToList();
                var isMove = content.RequestedOperation == DataPackageOperation.Move;
                var expectedPaths = sourcePaths
                    .Select(path => GetExpectedPastedPath(path, destPath, isMove))
                    .ToList();

                var operationCompletion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

                var operation = new Core.Models.FileOperation
                {
                    OperationType = isMove ? Core.Models.FileOperationType.Move : Core.Models.FileOperationType.Copy,
                    SourcePaths = sourcePaths,
                    DestinationPath = destPath
                };

                var fileService = App.Services.GetRequiredService<Core.Interfaces.IFileSystemService>();
                var progress = new Progress<Core.Models.FileOperationProgress>(p =>
                {
                    _dispatcherQueue.TryEnqueue(() =>
                    {
                        if (p.IsComplete)
                        {
                            OperationProgressBar.IsOpen = false;
                            operationCompletion.TrySetResult(true);
                        }
                        else if (p.IsFailed)
                        {
                            OperationProgressBar.IsOpen = false;
                            operationCompletion.TrySetResult(false);
                        }
                        else
                        {
                            OperationProgressBar.IsOpen = true;
                            OperationProgressBar.Message = $"Processing {p.CurrentFile}...";
                            OperationProgress.Value = p.ProgressPercent;
                        }
                    });
                });

                await fileService.EnqueueOperationAsync(operation, progress);

                var completed = await operationCompletion.Task;
                if (!completed)
                    return;

                if (ViewModel.ActiveTab is not null)
                {
                    await ViewModel.ActiveTab.RefreshAsync();
                    FileList.SelectAndRevealPaths(expectedPaths);
                }
            }
        }
        catch (Exception ex)
        {
            ShowError($"Paste failed: {ex.Message}");
        }
    }

    private static string GetExpectedPastedPath(string sourcePath, string destinationPath, bool isMove)
    {
        var directPath = Path.Combine(destinationPath, Path.GetFileName(sourcePath));
        if (isMove)
            return directPath;

        var sourceParent = Path.GetDirectoryName(sourcePath);
        var sameFolderCopy = string.Equals(
            NormalizeDirectoryPath(sourceParent),
            NormalizeDirectoryPath(destinationPath),
            StringComparison.OrdinalIgnoreCase);

        if (!sameFolderCopy)
            return directPath;

        return BuildExplorerCopyPath(directPath);
    }

    private static string BuildExplorerCopyPath(string existingPath)
    {
        var directory = Path.GetDirectoryName(existingPath) ?? string.Empty;
        var isDirectory = Directory.Exists(existingPath);
        var ext = isDirectory ? string.Empty : Path.GetExtension(existingPath);
        var baseName = isDirectory
            ? Path.GetFileName(existingPath)
            : Path.GetFileNameWithoutExtension(existingPath);

        var candidate = Path.Combine(directory, $"{baseName} - Copy{ext}");
        if (!File.Exists(candidate) && !Directory.Exists(candidate))
            return candidate;

        var counter = 2;
        while (true)
        {
            candidate = Path.Combine(directory, $"{baseName} - Copy ({counter}){ext}");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
                return candidate;

            counter++;
        }
    }

    private static string NormalizeDirectoryPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    #endregion

    #region Keyboard Accelerators

    private void RegisterKeyboardAccelerators()
    {
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.W, Windows.System.VirtualKeyModifiers.Control, (_, _) => { if (ViewModel.ActiveTab is not null) ViewModel.CloseTab(ViewModel.ActiveTab); }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Tab, Windows.System.VirtualKeyModifiers.Control, (_, _) => ViewModel.NextTab()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Tab, Windows.System.VirtualKeyModifiers.Control | Windows.System.VirtualKeyModifiers.Shift, (_, _) => ViewModel.PreviousTab()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.N, Windows.System.VirtualKeyModifiers.Control | Windows.System.VirtualKeyModifiers.Shift, (_, _) =>
        {
            if (ViewModel.ActiveTab is not null)
                _ = ViewModel.ActiveTab.CreateNewFolderAsync();
        }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.C, Windows.System.VirtualKeyModifiers.Control, (_, _) =>
        {
            if (IsTextInputFocused())
                return;

            CopyToClipboard();
        }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.X, Windows.System.VirtualKeyModifiers.Control, (_, _) =>
        {
            if (IsTextInputFocused())
                return;

            CutToClipboard();
        }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.V, Windows.System.VirtualKeyModifiers.Control, (_, _) =>
        {
            if (IsTextInputFocused())
                return;

            _ = PasteFromClipboardAsync();
        }));

        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Left, Windows.System.VirtualKeyModifiers.Menu, (_, _) => _ = ViewModel.GoBackAsync()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Right, Windows.System.VirtualKeyModifiers.Menu, (_, _) => _ = ViewModel.GoForwardAsync()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Back, Windows.System.VirtualKeyModifiers.None, (_, _) =>
        {
            if (IsTextInputFocused())
                return;

            if (!string.IsNullOrWhiteSpace(ViewModel.ActiveTab?.CurrentPath))
            {
                FileList.PrepareSelectionAfterGoUp(ViewModel.ActiveTab.CurrentPath);
            }

            _ = ViewModel.GoUpAsync();
        }));

        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F2, Windows.System.VirtualKeyModifiers.None, (_, _) => ViewModel.ActiveTab?.StartRename()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F3, Windows.System.VirtualKeyModifiers.None, (_, _) =>
        {
            PreviewColumn.Width = new GridLength(ViewModel.PreviewPaneWidth > 0 ? ViewModel.PreviewPaneWidth : 320);
        }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F5, Windows.System.VirtualKeyModifiers.None, (_, _) => _ = RefreshActiveTabAsync()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Delete, Windows.System.VirtualKeyModifiers.None, (_, _) =>
        {
            if (IsTextInputFocused())
                return;

            _ = DeleteSelectedWithConfirmationAsync();
        }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Delete, Windows.System.VirtualKeyModifiers.Shift, (_, _) => _ = HandleShiftDelete()));

        for (int i = 1; i <= 9; i++)
        {
            var idx = i;
            var key = (Windows.System.VirtualKey)(Windows.System.VirtualKey.Number1 + idx - 1);
            Content.KeyboardAccelerators.Add(CreateAccelerator(key, Windows.System.VirtualKeyModifiers.Control, (_, _) => ViewModel.GoToTab(idx)));
        }
    }

    private static KeyboardAccelerator CreateAccelerator(Windows.System.VirtualKey key, Windows.System.VirtualKeyModifiers modifiers, TypedEventHandler<KeyboardAccelerator, KeyboardAcceleratorInvokedEventArgs> handler)
    {
        var accel = new KeyboardAccelerator { Key = key, Modifiers = modifiers };
        accel.Invoked += (sender, args) =>
        {
            handler(sender, args);
            args.Handled = true;
        };
        return accel;
    }

    private bool IsTextInputFocused()
    {
        var focused = FocusManager.GetFocusedElement(Content.XamlRoot) as DependencyObject;
        while (focused is not null)
        {
            if (focused is TextBox or AutoSuggestBox or PasswordBox or RichEditBox)
                return true;

            focused = VisualTreeHelper.GetParent(focused);
        }

        return false;
    }

    private async Task HandleShiftDelete()
    {
        if (ViewModel.ActiveTab is null || ViewModel.ActiveTab.SelectedItems.Count == 0) return;

        var dialog = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            Title = "Permanently delete",
            Content = $"Are you sure you want to permanently delete {ViewModel.ActiveTab.SelectedItems.Count} item(s)? This action cannot be undone.",
            PrimaryButtonText = "Delete",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close
        };

        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            await ViewModel.ActiveTab.PermanentDeleteSelectedAsync();
        }
    }

    public async Task DeleteSelectedWithConfirmationAsync()
    {
        var activeTab = ViewModel.ActiveTab;
        if (activeTab is null || activeTab.SelectedItems.Count == 0)
            return;

        var hasNonEmptyFolder = activeTab.SelectedItems
            .Select(i => i.Entry.FullPath)
            .Where(Directory.Exists)
            .Any(DirectoryHasAnyContent);

        if (!hasNonEmptyFolder)
        {
            await activeTab.DeleteSelectedAsync();
            return;
        }

        var dialog = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            Title = "Delete non-empty folder(s)",
            Content = "One or more selected folders are not empty. Delete selected item(s) and move them to Recycle Bin?",
            PrimaryButtonText = "Delete",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close
        };

        var result = await dialog.ShowAsync();
        if (result == ContentDialogResult.Primary)
        {
            await activeTab.DeleteSelectedAsync();
        }
    }

    private static bool DirectoryHasAnyContent(string path)
    {
        try
        {
            return Directory.EnumerateFileSystemEntries(path).Any();
        }
        catch
        {
            // If we can't inspect safely, require confirmation to avoid silent destructive behavior.
            return true;
        }
    }

    #endregion

    #region Status Bar

    private void UpdateStatusBar()
    {
        if (ViewModel.ActiveTab is null) return;
        StatusBarLeft.Text = ViewModel.ActiveTab.StatusText;
        StatusBarRight.Text = ViewModel.ActiveTab.FreeSpaceText;
    }

    #endregion

    #region Error Display

    private void FileSystemService_OperationError(object? sender, FileOperationErrorEventArgs e)
    {
        if (!IsFileInUseError(e))
            return;

        _dispatcherQueue.TryEnqueue(() =>
        {
            _ = ShowFileInUseWarningAsync(e.FilePath);
        });
    }

    private async Task ShowFileInUseWarningAsync(string path)
    {
        if (_isFileInUseWarningOpen)
            return;

        _isFileInUseWarningOpen = true;
        try
        {
            var itemName = string.IsNullOrWhiteSpace(path) ? "item" : Path.GetFileName(path);
            var dialog = new ContentDialog
            {
                XamlRoot = Content.XamlRoot,
                Title = "File is open",
                Content = $"Cannot rename or delete '{itemName}' because it is open in another application. Close it and try again.",
                CloseButtonText = "OK",
                DefaultButton = ContentDialogButton.Close
            };

            await dialog.ShowAsync();
        }
        catch
        {
            // Ignore dialog display failures and keep the app responsive.
        }
        finally
        {
            _isFileInUseWarningOpen = false;
        }
    }

    private static bool IsFileInUseError(FileOperationErrorEventArgs e)
    {
        if (e.Win32ErrorCode is 32 or 33)
            return true;

        return e.ErrorMessage.Contains("being used by another process", StringComparison.OrdinalIgnoreCase) ||
               e.ErrorMessage.Contains("used by another process", StringComparison.OrdinalIgnoreCase);
    }

    private void ShowError(string message)
    {
        _dispatcherQueue.TryEnqueue(() =>
        {
            ErrorInfoBar.Message = message;
            ErrorInfoBar.IsOpen = true;
        });
    }

    #endregion

    #region Window Lifecycle

    private async void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        _fileSystemService.OperationError -= FileSystemService_OperationError;
        ViewModel.Tabs.CollectionChanged -= Tabs_CollectionChanged;

        foreach (var tab in _tabPropertySubscriptions.ToList())
        {
            tab.PropertyChanged -= OnTabPropertyChanged;
        }

        _tabPropertySubscriptions.Clear();
        _tabItemsById.Clear();

        if (_subscribedTab is not null)
        {
            _subscribedTab.Navigated -= OnActiveTabNavigated;
            _subscribedTab = null;
        }

        await ViewModel.SaveSessionAsync();
        ViewModel.Dispose();
    }

    #endregion
}
