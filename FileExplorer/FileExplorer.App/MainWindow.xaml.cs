// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Linq;
using FileExplorer.Core.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Input;
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
    private readonly DispatcherQueue _dispatcherQueue;

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
        ViewModel = App.Services.GetRequiredService<MainViewModel>();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(TabViewControl);

        RegisterKeyboardAccelerators();
        Closed += MainWindow_Closed;

        // Subscribe to tree sync so we can update the WinUI TreeView selection
        ViewModel.TreeView.SyncCompleted += OnTreeSyncCompleted;

        _ = InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        try
        {
            await ViewModel.InitializeAsync();

            UpdateTabBindings();
            UpdateStatusBar();

            PreviewPaneToggle.IsChecked = ViewModel.ShowPreviewPane;
            HiddenFilesToggle.IsChecked = ViewModel.ShowHiddenFiles;
            FileExtensionsToggle.IsChecked = ViewModel.ShowFileExtensions;

            TreeColumn.Width = new GridLength(ViewModel.NavigationPaneWidth > 0 ? ViewModel.NavigationPaneWidth : 315);

            // Populate tree — add TreeViewNode objects directly to RootNodes
            PopulateTreeView();

            if (ViewModel.ActiveTab is not null)
            {
                SubscribeToTabNavigation(ViewModel.ActiveTab);
                FileList.BindToTab(ViewModel.ActiveTab);
                UpdateAddressBar(ViewModel.ActiveTab.CurrentPath);
                UpdateSearchPlaceholder(ViewModel.ActiveTab.CurrentPath);
            }
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
        }
    }

    private void UpdateTabBindings()
    {
        TabViewControl.TabItemsSource = ViewModel.Tabs;
    }

    #endregion

    #region Navigation

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
            FileList.SetSearchMode(false);
            UpdateAddressBar(path);
            UpdateStatusBar();
            UpdateSearchPlaceholder(path);
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
        AddressBar.Text = ViewModel.ActiveTab?.CurrentPath ?? "";
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
        var folderName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar));
        SearchBox.PlaceholderText = string.IsNullOrEmpty(folderName) ? "Search..." : $"Search {folderName}";
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

    private async void NavigationTree_Expanding(TreeView sender, TreeViewExpandingEventArgs args)
    {
        // Skip if we're programmatically syncing the tree — nodes already have children
        if (_isSyncingTree) return;

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

    private void TreeSplitter_ManipulationDelta(object sender, ManipulationDeltaRoutedEventArgs e)
    {
        var newWidth = TreeColumn.Width.Value + e.Delta.Translation.X;
        if (newWidth >= 120 && newWidth <= 600)
        {
            TreeColumn.Width = new GridLength(newWidth);
            ViewModel.NavigationPaneWidth = newWidth;
        }
    }

    /// <summary>Flag to suppress the Expanding handler during programmatic sync.</summary>
    private bool _isSyncingTree;

    /// <summary>Called when TreeViewModel.SyncWithPathAsync finishes — syncs the WinUI TreeView control.</summary>
    private void OnTreeSyncCompleted(object? sender, TreeNodeViewModel selectedVm)
    {
        _dispatcherQueue.TryEnqueue(() =>
        {
            _isSyncingTree = true;
            try
            {
                // Full rebuild from authoritative VM state
                PopulateTreeView();

                // Expand nodes bottom-up: collect all nodes needing expansion, then expand deepest first
                var toExpand = new List<TreeViewNode>();
                CollectExpandedNodes(NavigationTree.RootNodes, toExpand);
                foreach (var tvNode in toExpand)
                {
                    tvNode.IsExpanded = true;
                }
            }
            finally
            {
                _isSyncingTree = false;
            }

            // Defer selection so WinUI can finish expanding layout
            _dispatcherQueue.TryEnqueue(() =>
            {
                var tvNode = FindTreeViewNode(NavigationTree.RootNodes, selectedVm);
                if (tvNode is not null)
                {
                    NavigationTree.SelectedNode = tvNode;
                }
            });
        });
    }

    /// <summary>Recursively collect TreeViewNodes whose ViewModel.IsExpanded is true, deepest-first.</summary>
    private static void CollectExpandedNodes(IList<TreeViewNode> nodes, List<TreeViewNode> result)
    {
        foreach (var tvNode in nodes)
        {
            // Recurse into children first (bottom-up)
            if (tvNode.Children.Count > 0)
                CollectExpandedNodes(tvNode.Children, result);

            if (tvNode.Content is TreeNodeViewModel vm && vm.IsExpanded)
                result.Add(tvNode);
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

    private void Delete_Click(object sender, RoutedEventArgs e) => _ = ViewModel.ActiveTab?.DeleteSelectedAsync();
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
                            _ = ViewModel.ActiveTab?.RefreshAsync();
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
            }
        }
        catch (Exception ex)
        {
            ShowError($"Paste failed: {ex.Message}");
        }
    }

    #endregion

    #region Keyboard Accelerators

    private void RegisterKeyboardAccelerators()
    {
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.T, Windows.System.VirtualKeyModifiers.Control, (_, _) => _ = ViewModel.NewTabAsync()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.W, Windows.System.VirtualKeyModifiers.Control, (_, _) => { if (ViewModel.ActiveTab is not null) ViewModel.CloseTab(ViewModel.ActiveTab); }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Tab, Windows.System.VirtualKeyModifiers.Control, (_, _) => ViewModel.NextTab()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Tab, Windows.System.VirtualKeyModifiers.Control | Windows.System.VirtualKeyModifiers.Shift, (_, _) => ViewModel.PreviousTab()));

        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Left, Windows.System.VirtualKeyModifiers.Menu, (_, _) => _ = ViewModel.GoBackAsync()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Right, Windows.System.VirtualKeyModifiers.Menu, (_, _) => _ = ViewModel.GoForwardAsync()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.Back, Windows.System.VirtualKeyModifiers.None, (_, _) => _ = ViewModel.GoUpAsync()));

        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F2, Windows.System.VirtualKeyModifiers.None, (_, _) => ViewModel.ActiveTab?.StartRename()));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F3, Windows.System.VirtualKeyModifiers.None, (_, _) => { ViewModel.TogglePreviewPane(); PreviewColumn.Width = ViewModel.ShowPreviewPane ? new GridLength(ViewModel.PreviewPaneWidth) : new GridLength(0); }));
        Content.KeyboardAccelerators.Add(CreateAccelerator(Windows.System.VirtualKey.F5, Windows.System.VirtualKeyModifiers.None, (_, _) => _ = ViewModel.RefreshAsync()));
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
        accel.Invoked += handler;
        return accel;
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
        ViewModel.TreeView.SyncCompleted -= OnTreeSyncCompleted;
        await ViewModel.SaveSessionAsync();
        ViewModel.Dispose();
    }

    #endregion
}
