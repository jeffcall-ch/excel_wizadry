// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using FileExplorer.Core.ViewModels;
using FileExplorer.Core.Models;
using FileExplorer.Core.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Dispatching;
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
    private const int IDC_ARROW = 32512;
    private const int IDC_SIZEWE = 32644;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern nint LoadCursor(nint hInstance, int lpCursorName);

    [DllImport("user32.dll")]
    private static extern nint SetCursor(nint hCursor);

    private readonly DispatcherQueue _dispatcherQueue;
    private readonly IFileSystemService _fileSystemService;
    private TreeNodeViewModel? _pendingSelectedTreeNode;
    private string? _lastSyncedTreePath;
    private readonly string _shortcutsFolderPath;
    private bool _isFileInUseWarningOpen;
    private readonly object _shortcutCacheLock = new();
    private readonly Dictionary<string, ShortcutFolderCacheData> _shortcutFolderCache = new(StringComparer.OrdinalIgnoreCase);
    private List<ShortcutRootEntryData> _shortcutRootCache = [];
    private MenuFlyout? _shortcutLauncherFlyout;
    private bool _shortcutPreloadStarted;
    private const int ShortcutMenuMaxItems = 50;
    private const int ShortcutMenuMaxSubfolders = 50;
    private const int ShortcutMenuMaxDepth = 64;

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
    }

    /// <summary>Gets the main view model.</summary>
    public MainViewModel ViewModel { get; }

    /// <summary>Gets the native window handle.</summary>
    public nint Hwnd { get; }

    /// <summary>Initializes a new instance of the <see cref="MainWindow"/> class.</summary>
    public MainWindow()
    {
        InitializeComponent();

        Title = "File Explorer";
        SystemBackdrop = new MicaBackdrop();

        Hwnd = WindowNative.GetWindowHandle(this);
        _dispatcherQueue = DispatcherQueue.GetForCurrentThread();
        _fileSystemService = App.Services.GetRequiredService<IFileSystemService>();
        ViewModel = App.Services.GetRequiredService<MainViewModel>();
        _shortcutsFolderPath = ResolveShortcutsFolderPath();
        _fileSystemService.OperationError += FileSystemService_OperationError;

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(TabViewControl);

        RegisterKeyboardAccelerators();
        Closed += MainWindow_Closed;
        NavigationTree.Loaded += NavigationTree_Loaded;

        StartShortcutCachePreload();

        _ = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        try
        {
            await ViewModel.InitializeAsync(App.StartupPath);

            TreeNodeViewModel? initialSelectedTreeNode = null;
            if (ViewModel.ActiveTab is not null)
            {
                initialSelectedTreeNode = await ViewModel.TreeView.SyncWithPathAsync(ViewModel.ActiveTab.CurrentPath);
                _lastSyncedTreePath = NormalizeTreeSyncPath(ViewModel.ActiveTab.CurrentPath);
            }

            UpdateTabBindings();
            UpdateStatusBar();

            TreeColumn.Width = new GridLength(ViewModel.NavigationPaneWidth > 0 ? ViewModel.NavigationPaneWidth : 315);

            // Populate tree — add TreeViewNode objects directly to RootNodes
            PopulateTreeView();

            if (ViewModel.ActiveTab is not null)
            {
                SubscribeToTabNavigation(ViewModel.ActiveTab);
                FileList.BindToTab(ViewModel.ActiveTab);
                UpdateAddressBar(ViewModel.ActiveTab.CurrentPath);
                UpdateSearchPlaceholder(ViewModel.ActiveTab.CurrentPath);

                if (_navigationTreeLoaded)
                {
                    ApplyTreeSelection(initialSelectedTreeNode);
                }
                else
                {
                    _pendingSelectedTreeNode = initialSelectedTreeNode;
                }
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
        _ = ViewModel.NewTabAsync();
        UpdateTabBindings();
    }

    private void TabView_TabCloseRequested(TabView sender, TabViewTabCloseRequestedEventArgs args)
    {
        if (args.Item is TabContentViewModel tab)
        {
            ViewModel.CloseTab(tab);
            UpdateTabBindings();
        }
    }

    private void TabView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (TabViewControl.SelectedItem is TabContentViewModel tab)
        {
            ViewModel.ActiveTab = tab;
            SubscribeToTabNavigation(tab);
            FileList.BindToTab(tab);
            UpdateAddressBar(tab.CurrentPath);
            UpdateStatusBar();
            UpdateSearchPlaceholder(tab.CurrentPath);
            _ = SyncNavigationTreeAsync(tab.CurrentPath);
        }
    }

    private void UpdateTabBindings()
    {
        TabViewControl.TabItemsSource = ViewModel.Tabs;
        TabViewControl.SelectedItem = ViewModel.ActiveTab;
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
        }
    }

    private TabContentViewModel? _subscribedTab;

    private void BackButton_Click(object sender, RoutedEventArgs e) => _ = ViewModel.GoBackAsync();
    private void ForwardButton_Click(object sender, RoutedEventArgs e) => _ = ViewModel.GoForwardAsync();
    private void UpButton_Click(object sender, RoutedEventArgs e) => _ = ViewModel.GoUpAsync();
    private void RefreshButton_Click(object sender, RoutedEventArgs e) => _ = ViewModel.RefreshAsync();

    private void SubscribeToTabNavigation(TabContentViewModel? tab)
    {
        if (_subscribedTab is not null)
            _subscribedTab.Navigated -= OnActiveTabNavigated;

        _subscribedTab = tab;

        if (tab is not null)
            tab.Navigated += OnActiveTabNavigated;
    }

    private void OnActiveTabNavigated(object? sender, string path)
    {
        _dispatcherQueue.TryEnqueue(() =>
        {
            UpdateTabBindings();
            FileList.SetSearchMode(false);
            UpdateAddressBar(path);
            UpdateStatusBar();
            UpdateSearchPlaceholder(path);
            _ = SyncNavigationTreeAsync(path);
        });
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
                var chevron = new FontIcon
                {
                    FontFamily = new Microsoft.UI.Xaml.Media.FontFamily("Segoe Fluent Icons"),
                    Glyph = "\uE76C",
                    FontSize = 10,
                    Opacity = 0.6
                };
                chevron.VerticalAlignment = VerticalAlignment.Center;
                chevron.Margin = new Thickness(0);
                BreadcrumbPanel.Children.Add(chevron);
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

    private void ShortcutLauncherButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StartShortcutCachePreload();
            var flyout = GetOrCreateShortcutLauncherFlyout();
            flyout.ShowAt(ShortcutLauncherButton);
        }
        catch (Exception ex)
        {
            ShowError($"Failed to open shortcuts launcher: {ex.Message}");
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
            flyout.Items.Add(new MenuFlyoutItem
            {
                Text = "Shortcuts folder not found",
                IsEnabled = false
            });
            return flyout;
        }

        var entries = GetShortcutRootEntriesForFlyout();
        if (entries.Count == 0)
        {
            flyout.Items.Add(new MenuFlyoutItem
            {
                Text = "No shortcuts found",
                IsEnabled = false
            });
        }
        else
        {
            foreach (var entry in entries)
            {
                AddShortcutRootEntry(flyout, entry);
            }
        }

        flyout.Items.Add(new MenuFlyoutSeparator());

        var openShortcutsFolder = new MenuFlyoutItem
        {
            Text = "Open shortcuts folder in new tab",
            Icon = new FontIcon { Glyph = "\uE8B7" }
        };
        openShortcutsFolder.Click += (_, _) => _ = OpenFolderInNewTabAsync(_shortcutsFolderPath);
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

    private async Task PreloadShortcutCacheAsync()
    {
        try
        {
            var rootEntries = LoadShortcutRootEntries();

            // Phase 1: quickly cache root folders' immediate children so all first-layer submenus are ready.
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

            _dispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, () =>
            {
                _ = GetOrCreateShortcutLauncherFlyout();
            });

            await Task.Yield();

            // Phase 2: silently pre-populate every reachable layer in the background.
            var deepEntries = new Dictionary<string, ShortcutFolderCacheData>(StringComparer.OrdinalIgnoreCase);
            foreach (var root in rootEntries)
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
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Shortcut cache preload failed: {ex.Message}");
        }
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
            var item = new MenuFlyoutItem { Text = entry.DisplayName };
            ToolTipService.SetToolTip(item, entry.ShortcutPath);
            item.Click += (_, _) => OpenShortcutTarget(entry.ShortcutPath);
            flyout.Items.Add(item);
        }
    }

    private MenuFlyoutSubItem CreateFolderSubMenu(string displayName, string folderPath, int depth)
    {
        var folderItem = new MenuFlyoutSubItem { Text = displayName };
        ToolTipService.SetToolTip(folderItem, folderPath);
        var context = new ShortcutFolderMenuContext(folderPath, depth);
        folderItem.Tag = context;
        folderItem.AddHandler(UIElement.PointerPressedEvent, new PointerEventHandler(OnShortcutFolderItemPointerPressed), true);

        if (depth >= ShortcutMenuMaxDepth)
            return folderItem;

        if (TryGetShortcutFolderCache(folderPath, out var cachedData))
        {
            var hasCachedChildren = cachedData.Subfolders.Count > 0 || cachedData.Files.Count > 0;
            if (!hasCachedChildren)
                return folderItem;

            if (depth == 0)
            {
                // Requirement: first layer is already populated when flyout opens.
                AddFolderSubMenuItems(folderItem, cachedData, depth + 1);
                context.IsLoaded = true;
            }
            else
            {
                InitializeFolderLazyLoading(folderItem);
            }

            return folderItem;
        }

        var hasChildren = GetSubfolders(folderPath, 1).Count > 0 || GetFiles(folderPath, 1).Count > 0;
        if (hasChildren)
        {
            InitializeFolderLazyLoading(folderItem);
        }

        return folderItem;
    }

    private void InitializeFolderLazyLoading(MenuFlyoutSubItem folderItem)
    {
        folderItem.Items.Add(new MenuFlyoutItem { Text = string.Empty, IsEnabled = false });
        folderItem.PointerEntered += (_, _) => EnsureFolderSubMenuLoaded(folderItem);
        folderItem.GotFocus += (_, _) => EnsureFolderSubMenuLoaded(folderItem);
        folderItem.KeyDown += (_, _) => EnsureFolderSubMenuLoaded(folderItem);
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

        _ = OpenFolderInNewTabAsync(context.FolderPath);
        ShortcutLauncherButton.Flyout?.Hide();

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

    private void EnsureFolderSubMenuLoaded(MenuFlyoutSubItem folderItem)
    {
        if (folderItem.Tag is not ShortcutFolderMenuContext context || context.IsLoaded)
            return;

        context.IsLoaded = true;
        folderItem.Items.Clear();

        if (context.Depth >= ShortcutMenuMaxDepth)
            return;

        if (TryGetShortcutFolderCache(context.FolderPath, out var cachedData))
        {
            AddFolderSubMenuItems(folderItem, cachedData, context.Depth + 1);
            return;
        }

        var subfolders = GetSubfolders(context.FolderPath, ShortcutMenuMaxSubfolders);
        var files = GetFiles(context.FolderPath, ShortcutMenuMaxSubfolders);

        var fallbackData = new ShortcutFolderCacheData(
            [.. subfolders.Select(sub => new ShortcutSubfolderData(sub.Name, sub.FullName))],
            [.. files.Select(file => new ShortcutFileData(file.Name, file.FullName))],
            IsFullyLoaded: false);

        lock (_shortcutCacheLock)
        {
            _shortcutFolderCache[context.FolderPath] = fallbackData;
        }

        AddFolderSubMenuItems(folderItem, fallbackData, context.Depth + 1);
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
            var fileItem = new MenuFlyoutItem { Text = file.Name };
            ToolTipService.SetToolTip(fileItem, file.FullPath);
            fileItem.Click += (_, _) => OpenFileOrShortcut(file.FullPath);
            folderItem.Items.Add(fileItem);
        }
    }

    private bool TryGetShortcutFolderCache(string folderPath, out ShortcutFolderCacheData data)
    {
        lock (_shortcutCacheLock)
        {
            return _shortcutFolderCache.TryGetValue(folderPath, out data!);
        }
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
            if (Directory.Exists(path))
            {
                _ = OpenFolderInNewTabAsync(path);
                return;
            }

            var shellService = App.Services.GetRequiredService<IShellIntegrationService>();
            shellService.OpenFile(path);
        }
        catch (Exception ex)
        {
            ShowError($"Failed to open item: {ex.Message}");
        }
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
        var newWidth = TreeColumn.Width.Value + e.Delta.Translation.X;
        if (newWidth >= 120 && newWidth <= 600)
        {
            TreeColumn.Width = new GridLength(newWidth);
            ViewModel.NavigationPaneWidth = newWidth;
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
            TabViewControl.SelectedItem = ViewModel.ActiveTab;
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
        flyout.Items.Add(new MenuFlyoutItem { Text = "Open", Icon = new FontIcon { Glyph = "\uE8E5" } });
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            _ = ViewModel.NavigateToAddressAsync(path);
        };

        flyout.Items.Add(new MenuFlyoutItem { Text = "Open in new tab", Icon = new FontIcon { Glyph = "\uE7C3" } });
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            _ = OpenFolderInNewTabAsync(path);
        };

        flyout.Items.Add(new MenuFlyoutItem { Text = "Open in new window", Icon = new FontIcon { Glyph = "\uE8A7" } });
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            OpenFolderInNewWindow(path);
        };
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
        ViewModel.TogglePreviewPane();
        PreviewColumn.Width = ViewModel.ShowPreviewPane ? new GridLength(ViewModel.PreviewPaneWidth) : new GridLength(0);
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

            _ = ViewModel.GoUpAsync();
        }));

        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F2, Windows.System.VirtualKeyModifiers.None, (_, _) => ViewModel.ActiveTab?.StartRename()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F3, Windows.System.VirtualKeyModifiers.None, (_, _) => { ViewModel.TogglePreviewPane(); PreviewColumn.Width = ViewModel.ShowPreviewPane ? new GridLength(ViewModel.PreviewPaneWidth) : new GridLength(0); }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F5, Windows.System.VirtualKeyModifiers.None, (_, _) => _ = ViewModel.RefreshAsync()));
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
        await ViewModel.SaveSessionAsync();
        ViewModel.Dispose();
    }

    #endregion
}
