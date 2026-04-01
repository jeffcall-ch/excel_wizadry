// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;

namespace FileExplorer.App.Views;

/// <summary>
/// File list details view with virtualized items, column headers with sort, inline rename,
/// drag-drop support, and OneDrive status column.
/// </summary>
public sealed partial class FileListView : UserControl
{
    private TabContentViewModel? _currentTab;
    private string _typeAheadBuffer = string.Empty;
    private DateTimeOffset _lastKeyTime;
    private bool _isSearchMode;
    private bool _isClearingSelection;
    private int _selectionAnchorIndex = -1;
    private bool _isProgrammaticRangeSelection;
    private bool _isShiftPressed;
    private bool _restoreFocusAfterNavigation;
    private string? _pendingSelectPathAfterGoUp;

    /// <summary>Initializes a new instance of the <see cref="FileListView"/> class.</summary>
    public FileListView()
    {
        InitializeComponent();
    }

    public void FocusForKeyboardNavigation()
    {
        DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            FileListView_Inner.Focus(FocusState.Programmatic);
        });
    }

    public void PrepareSelectionAfterGoUp(string? previouslyOpenedFolderPath)
    {
        if (!string.IsNullOrWhiteSpace(previouslyOpenedFolderPath))
        {
            _pendingSelectPathAfterGoUp = previouslyOpenedFolderPath;
        }
    }

    /// <summary>Binds this view to a tab's content ViewModel.</summary>
    public void BindToTab(TabContentViewModel tab)
    {
        if (_currentTab is not null)
        {
            _currentTab.Items.CollectionChanged -= CurrentTab_ItemsCollectionChanged;
            _currentTab.PropertyChanged -= CurrentTab_PropertyChanged;
            _currentTab.RevealItemRequested -= CurrentTab_RevealItemRequested;
        }

        _currentTab = tab;
        FileListView_Inner.ItemsSource = tab.Items;
        ClearTransientSelection();
        UpdateEmptyState();
        UpdateSortIndicators();
        UpdateOneDriveColumnVisibility();

        tab.Items.CollectionChanged += CurrentTab_ItemsCollectionChanged;
        tab.PropertyChanged += CurrentTab_PropertyChanged;
        tab.RevealItemRequested += CurrentTab_RevealItemRequested;
    }

    private void CurrentTab_ItemsCollectionChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        DispatcherQueue.TryEnqueue(() =>
        {
            UpdateEmptyState();

            TryApplyPendingSelectionAfterGoUp();

            if (e.NewItems is null) return;

            foreach (var item in e.NewItems.OfType<FileItemViewModel>())
            {
                if (item.IsRenaming || item.IsNewItem)
                {
                    RevealItem(item);
                }
            }
        });
    }

    private void CurrentTab_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (_currentTab is null) return;

        if (e.PropertyName == nameof(TabContentViewModel.IsLoading))
        {
            DispatcherQueue.TryEnqueue(() =>
            {
                if (_currentTab.IsLoading)
                {
                    ClearTransientSelection();
                }

                LoadingRing.Visibility = _currentTab.IsLoading ? Visibility.Visible : Visibility.Collapsed;
                LoadingRing.IsActive = _currentTab.IsLoading;

                if (!_currentTab.IsLoading && _restoreFocusAfterNavigation)
                {
                    _restoreFocusAfterNavigation = false;
                    DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
                    {
                        FileListView_Inner.Focus(FocusState.Programmatic);
                    });
                }

                if (!_currentTab.IsLoading)
                {
                    if (!TryApplyPendingSelectionAfterGoUp())
                    {
                        EnsureTopRowVisibleWhenNoSelection();
                    }
                }
            });
        }
    }

    private void CurrentTab_RevealItemRequested(object? sender, FileItemViewModel item)
    {
        _typeAheadBuffer = string.Empty;
        DispatcherQueue.TryEnqueue(() => RevealItem(item));
    }

    #region Column Headers

    private void ColumnHeader_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string column && _currentTab is not null)
        {
            _currentTab.SetSortCommand.Execute(column);
            UpdateSortIndicators();
        }
    }

    private void ColumnHeader_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (sender is FrameworkElement element)
        {
            FlyoutBase.GetAttachedFlyout(ColumnHeaderGrid)?.ShowAt(element);
        }
    }

    private void ColumnVisibility_Changed(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox cb && cb.Tag is string column)
        {
            bool visible = cb.IsChecked == true;
            switch (column)
            {
                case "OneDrive":
                    ColOneDrive.Width = visible ? new GridLength(60) : new GridLength(0);
                    break;
                case "DateModified":
                    ColDateModified.Width = visible ? new GridLength(160) : new GridLength(0);
                    break;
                case "Type":
                    ColType.Width = visible ? new GridLength(130) : new GridLength(0);
                    break;
                case "Size":
                    ColSize.Width = visible ? new GridLength(80) : new GridLength(0);
                    break;
            }
        }
    }

    private void UpdateSortIndicators()
    {
        if (_currentTab is null) return;

        NameSortIcon.Visibility = Visibility.Collapsed;
        DateSortIcon.Visibility = Visibility.Collapsed;
        TypeSortIcon.Visibility = Visibility.Collapsed;
        SizeSortIcon.Visibility = Visibility.Collapsed;

        var upGlyph = "\uE70E";   // ChevronUp
        var downGlyph = "\uE70D"; // ChevronDown
        var glyph = _currentTab.SortAscending ? upGlyph : downGlyph;

        switch (_currentTab.SortColumn)
        {
            case "Name":
                NameSortIcon.Glyph = glyph;
                NameSortIcon.Visibility = Visibility.Visible;
                break;
            case "DateModified":
                DateSortIcon.Glyph = glyph;
                DateSortIcon.Visibility = Visibility.Visible;
                break;
            case "Type":
                TypeSortIcon.Glyph = glyph;
                TypeSortIcon.Visibility = Visibility.Visible;
                break;
            case "Size":
                SizeSortIcon.Glyph = glyph;
                SizeSortIcon.Visibility = Visibility.Visible;
                break;
        }
    }

    private void UpdateOneDriveColumnVisibility()
    {
        var mainVm = App.Services.GetRequiredService<MainViewModel>();
        if (!mainVm.ShowOneDriveColumn)
        {
            ColOneDrive.Width = new GridLength(0);
            OneDriveHeader.Visibility = Visibility.Collapsed;
            return;
        }

        ColOneDrive.Width = new GridLength(60);
        OneDriveHeader.Visibility = Visibility.Visible;
        ColOneDriveCheck.IsChecked = true;
    }

    #endregion

    #region Item Interactions

    private void FileList_ItemClick(object sender, ItemClickEventArgs e)
    {
        // Single click selects (handled by ListView default behavior)
    }

    private void RightPane_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (FindAncestorOfType<TextBox>(e.OriginalSource as DependencyObject) is not null)
            return;

        if (FileListView_Inner.FocusState == FocusState.Unfocused)
        {
            FileListView_Inner.Focus(FocusState.Programmatic);
        }
    }

    private void FileList_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (_currentTab is null) return;
        var item = ResolveItemFromOriginalSource(e.OriginalSource as DependencyObject) ?? GetActiveItem();
        if (item is not null)
        {
            FileListView_Inner.SelectedItem = item;
            _ = OpenItemFromListAsync(item);
        }
    }

    private async void FileListView_Inner_PreviewKeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (_currentTab is null)
            return;

        if (IsShiftKey(e.Key))
        {
            _isShiftPressed = true;
            return;
        }

        if (_currentTab.Items.Any(item => item.IsRenaming))
            return;

        switch (e.Key)
        {
            case Windows.System.VirtualKey.Enter:
                if (GetActiveItem() is { } enterItem)
                {
                    e.Handled = true;
                    await OpenItemFromListAsync(enterItem);
                }
                break;

            case Windows.System.VirtualKey.Right:
                if (GetActiveItem() is { } rightItem && CanOpenWithRightArrow(rightItem))
                {
                    e.Handled = true;
                    await OpenItemFromListAsync(rightItem);
                }
                break;

            case Windows.System.VirtualKey.Left:
                if (!string.IsNullOrWhiteSpace(_currentTab.CurrentPath))
                {
                    _pendingSelectPathAfterGoUp = _currentTab.CurrentPath;
                    e.Handled = true;
                    _restoreFocusAfterNavigation = true;
                    ClearTransientSelection();
                    await _currentTab.GoUpAsync();
                    EnsurePendingSelectionAfterGoUpAsync();
                }
                break;

            case Windows.System.VirtualKey.Escape:
                e.Handled = true;
                ClearTransientSelection();
                FileListView_Inner.Focus(FocusState.Programmatic);
                break;

            case Windows.System.VirtualKey.Up:
                if (!IsShiftSelectionActive() && FileListView_Inner.SelectedItem is null && _currentTab.Items.Count > 0)
                {
                    e.Handled = true;
                    var last = _currentTab.Items[^1];
                    FileListView_Inner.SelectedItem = last;
                    _selectionAnchorIndex = _currentTab.Items.Count - 1;
                    FileListView_Inner.ScrollIntoView(last, ScrollIntoViewAlignment.Leading);
                }
                break;

            case Windows.System.VirtualKey.Down:
                // Preserve native ListView range selection behavior (Shift+Arrow).
                if (!IsShiftSelectionActive() && FileListView_Inner.SelectedItem is null && _currentTab.Items.Count > 0)
                {
                    e.Handled = true;
                    var first = _currentTab.Items[0];
                    FileListView_Inner.SelectedItem = first;
                    _selectionAnchorIndex = 0;
                    FileListView_Inner.ScrollIntoView(first, ScrollIntoViewAlignment.Leading);
                }
                break;
        }
    }

    private void FileListView_Inner_KeyUp(object sender, KeyRoutedEventArgs e)
    {
        if (IsShiftKey(e.Key))
        {
            _isShiftPressed = false;
        }
    }

    private void FileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_currentTab is null) return;

        if (_isClearingSelection)
            return;

        _currentTab.SelectedItems.Clear();
        foreach (var item in FileListView_Inner.SelectedItems.OfType<FileItemViewModel>())
        {
            _currentTab.SelectedItems.Add(item);
        }

        if (!_isProgrammaticRangeSelection && FileListView_Inner.SelectedItem is FileItemViewModel selected)
        {
            _selectionAnchorIndex = _currentTab.Items.IndexOf(selected);
        }

        ApplySelectionFrames();
    }

    /// <summary>Sets whether the list is showing search results.</summary>
    public void SetSearchMode(bool isSearch) => _isSearchMode = isSearch;

    /// <summary>Selects and reveals items whose full paths match the provided paths.</summary>
    public void SelectAndRevealPaths(IReadOnlyList<string> fullPaths)
    {
        if (_currentTab is null || fullPaths.Count == 0)
            return;

        var pathSet = new HashSet<string>(fullPaths, StringComparer.OrdinalIgnoreCase);
        var matches = _currentTab.Items
            .Where(item => pathSet.Contains(item.Entry.FullPath))
            .ToList();

        if (matches.Count == 0)
            return;

        FileListView_Inner.SelectedItems.Clear();
        foreach (var match in matches)
        {
            FileListView_Inner.SelectedItems.Add(match);
        }

        var first = matches[0];
        FileListView_Inner.SelectedItem = first;
        FileListView_Inner.ScrollIntoView(first, ScrollIntoViewAlignment.Leading);
        FileListView_Inner.Focus(FocusState.Programmatic);
    }

    private void FileList_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        try
        {
            if (_currentTab is null) return;

            // Get the item under the pointer
            FileItemViewModel? tappedItem = null;
            if (e.OriginalSource is FrameworkElement element && element.DataContext is FileItemViewModel item)
            {
                tappedItem = item;
                if (!FileListView_Inner.SelectedItems.Contains(item))
                {
                    FileListView_Inner.SelectedItem = item;
                }
            }

            if (_currentTab.SelectedItems.Count == 0 && tappedItem is null) return;

            var contextPaths = GetContextMenuPaths(tappedItem);

            var flyout = new MenuFlyout();

            // Standard items
            flyout.Items.Add(new MenuFlyoutItem { Text = "Open", Icon = new FontIcon { Glyph = "\uE8E5" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                var sel = _currentTab.SelectedItems.FirstOrDefault();
                if (sel is not null) _ = _currentTab.OpenItemAsync(sel);
            };

            var folderTarget = ResolveFolderTarget(tappedItem);
            if (folderTarget is not null && App.MainWindow is MainWindow mainWindow)
            {
                flyout.Items.Add(new MenuFlyoutItem { Text = "Open in new tab", Icon = new FontIcon { Glyph = "\uE7C3" } });
                ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
                {
                    flyout.Hide();
                    DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
                    {
                        _ = mainWindow.OpenFolderInNewTabAsync(folderTarget.Entry.FullPath);
                    });
                };

                flyout.Items.Add(new MenuFlyoutItem { Text = "Open in new window", Icon = new FontIcon { Glyph = "\uE8A7" } });
                ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
                {
                    flyout.Hide();
                    DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
                    {
                        mainWindow.OpenFolderInNewWindow(folderTarget.Entry.FullPath);
                    });
                };
            }

            if (_isSearchMode)
            {
                flyout.Items.Add(new MenuFlyoutItem { Text = "Open file location", Icon = new FontIcon { Glyph = "\uE8DA" } });
                ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
                {
                    var sel = _currentTab.SelectedItems.FirstOrDefault() ?? tappedItem;
                    if (sel?.Entry.ParentPath is string parentPath)
                    {
                        var mainVm = App.Services.GetRequiredService<MainViewModel>();
                        _ = mainVm.NavigateToAddressAsync(parentPath);
                    }
                };
            }

            flyout.Items.Add(new MenuFlyoutSeparator());

            flyout.Items.Add(new MenuFlyoutItem { Text = "Cut", Icon = new FontIcon { Glyph = "\uE8C6" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                if (App.MainWindow is MainWindow mw) mw.CutToClipboard();
            };

            flyout.Items.Add(new MenuFlyoutItem { Text = "Copy", Icon = new FontIcon { Glyph = "\uE8C8" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                if (App.MainWindow is MainWindow mw) mw.CopyToClipboard();
            };

            flyout.Items.Add(new MenuFlyoutItem { Text = "Delete", Icon = new FontIcon { Glyph = "\uE74D" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                if (App.MainWindow is MainWindow mw)
                {
                    _ = mw.DeleteSelectedWithConfirmationAsync();
                }
                else
                {
                    _ = _currentTab.DeleteSelectedAsync();
                }
            };

            flyout.Items.Add(new MenuFlyoutItem { Text = "Rename", Icon = new FontIcon { Glyph = "\uE8AC" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                _currentTab.StartRename();
            };

            flyout.Items.Add(new MenuFlyoutSeparator());

            flyout.Items.Add(new MenuFlyoutItem { Text = "Copy path", Icon = new FontIcon { Glyph = "\uE71B" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                var sel = _currentTab.SelectedItems.FirstOrDefault() ?? tappedItem;
                if (sel is not null)
                {
                    var dp = new DataPackage();
                    dp.SetText(sel.Entry.FullPath);
                    Clipboard.SetContent(dp);
                }
            };

            flyout.Items.Add(new MenuFlyoutItem { Text = "Properties", Icon = new FontIcon { Glyph = "\uE946" } });
            ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
            {
                if (App.MainWindow is MainWindow mw && contextPaths.Count > 0)
                {
                    var shellService = App.Services.GetRequiredService<Core.Interfaces.IShellIntegrationService>();
                    shellService.ShowProperties(mw.Hwnd, contextPaths);
                }
            };

            flyout.ShowAt(FileListView_Inner, e.GetPosition(FileListView_Inner));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"FileList right-click menu failed: {ex}");
        }
    }

    #endregion

    #region Inline Rename

    private void RenameTextBox_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (sender is not TextBox tb || tb.DataContext is not FileItemViewModel item) return;

        if (e.Key == Windows.System.VirtualKey.Enter)
        {
            e.Handled = true;
            item.EditingName = tb.Text;
            if (_currentTab is not null)
            {
                _ = _currentTab.CommitRenameAsync(item);
            }
        }
        else if (e.Key == Windows.System.VirtualKey.Escape)
        {
            e.Handled = true;
            item.IsRenaming = false;
            item.EditingName = item.Name;

            DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
            {
                if (FileListView_Inner.SelectedItem is null)
                {
                    FileListView_Inner.SelectedItem = item;
                }

                FileListView_Inner.Focus(FocusState.Programmatic);
            });
        }
    }

    private void RenameTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox tb && tb.DataContext is FileItemViewModel item && item.IsRenaming)
        {
            item.EditingName = tb.Text;
            if (_currentTab is not null)
            {
                _ = _currentTab.CommitRenameAsync(item);
            }
        }
    }

    private void RenameTextBox_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox tb)
        {
            tb.Focus(FocusState.Programmatic);
            // Select filename without extension
            var dotIndex = tb.Text.LastIndexOf('.');
            if (dotIndex > 0)
                tb.Select(0, dotIndex);
            else
                tb.SelectAll();
        }
    }

    private void RevealItem(FileItemViewModel item)
    {
        FileListView_Inner.SelectedItem = item;
        FileListView_Inner.ScrollIntoView(item, ScrollIntoViewAlignment.Leading);

        DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            FileListView_Inner.UpdateLayout();
            if (FileListView_Inner.ContainerFromItem(item) is ListViewItem container)
            {
                container.StartBringIntoView();

                if (item.IsRenaming && FindChildOfType<TextBox>(container) is TextBox tb)
                {
                    tb.Focus(FocusState.Programmatic);
                    var dotIndex = tb.Text.LastIndexOf('.');
                    if (dotIndex > 0)
                        tb.Select(0, dotIndex);
                    else
                        tb.SelectAll();
                    return;
                }
            }

            // After inline rename commit (e.g., new folder), restore list focus so Enter/Arrows work immediately.
            FileListView_Inner.Focus(FocusState.Programmatic);
        });
    }

    private async Task OpenItemFromListAsync(FileItemViewModel item)
    {
        if (item.Entry.IsDirectory || string.Equals(item.Entry.Extension, ".lnk", StringComparison.OrdinalIgnoreCase))
        {
            _restoreFocusAfterNavigation = true;
            ClearTransientSelection();
        }

        await _currentTab!.OpenItemAsync(item);
    }

    private FileItemViewModel? ResolveItemFromOriginalSource(DependencyObject? source)
    {
        while (source is not null)
        {
            if (source is FrameworkElement element && element.DataContext is FileItemViewModel item)
                return item;

            source = VisualTreeHelper.GetParent(source);
        }

        return null;
    }

    private FileItemViewModel? GetActiveItem()
    {
        if (FileListView_Inner.SelectedItem is FileItemViewModel selectedItem)
            return selectedItem;

        return _currentTab?.Items.FirstOrDefault();
    }

    private bool MoveSelection(int delta, bool extendRange)
    {
        if (_currentTab is null || _currentTab.Items.Count == 0)
            return false;

        var currentIndex = FileListView_Inner.SelectedItem is FileItemViewModel selectedItem
            ? _currentTab.Items.IndexOf(selectedItem)
            : -1;

        var nextIndex = currentIndex < 0
            ? (delta > 0 ? 0 : _currentTab.Items.Count - 1)
            : Math.Clamp(currentIndex + delta, 0, _currentTab.Items.Count - 1);

        if (nextIndex < 0 || nextIndex >= _currentTab.Items.Count)
            return false;

        var nextItem = _currentTab.Items[nextIndex];

        if (extendRange)
        {
            if (_selectionAnchorIndex < 0)
                _selectionAnchorIndex = currentIndex >= 0 ? currentIndex : nextIndex;

            ApplyRangeSelection(_selectionAnchorIndex, nextIndex);
        }
        else
        {
            _selectionAnchorIndex = nextIndex;
            FileListView_Inner.SelectedItems.Clear();
            FileListView_Inner.SelectedItem = nextItem;
        }

        FileListView_Inner.ScrollIntoView(nextItem);
        return true;
    }

    private void ApplyRangeSelection(int anchorIndex, int targetIndex)
    {
        if (_currentTab is null || _currentTab.Items.Count == 0)
            return;

        var start = Math.Max(0, Math.Min(anchorIndex, targetIndex));
        var end = Math.Min(_currentTab.Items.Count - 1, Math.Max(anchorIndex, targetIndex));

        _isProgrammaticRangeSelection = true;
        try
        {
            FileListView_Inner.SelectedItems.Clear();
            for (var i = start; i <= end; i++)
            {
                FileListView_Inner.SelectedItems.Add(_currentTab.Items[i]);
            }

            FileListView_Inner.SelectedItem = _currentTab.Items[targetIndex];
        }
        finally
        {
            _isProgrammaticRangeSelection = false;
        }
    }

    private bool IsShiftSelectionActive()
    {
        if (_isShiftPressed)
            return true;

        var left = InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.LeftShift);
        var right = InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.RightShift);
        var generic = InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift);

        return ((left | right | generic) & Windows.UI.Core.CoreVirtualKeyStates.Down) == Windows.UI.Core.CoreVirtualKeyStates.Down;
    }

    private static bool IsShiftKey(Windows.System.VirtualKey key)
    {
        return key == Windows.System.VirtualKey.Shift ||
               key == Windows.System.VirtualKey.LeftShift ||
               key == Windows.System.VirtualKey.RightShift;
    }

    private IReadOnlyList<string> GetContextMenuPaths(FileItemViewModel? tappedItem)
    {
        if (_currentTab is null)
            return [];

        if (tappedItem is not null && _currentTab.SelectedItems.Contains(tappedItem))
            return _currentTab.SelectedItems.Select(item => item.Entry.FullPath).ToList();

        if (tappedItem is not null)
            return [tappedItem.Entry.FullPath];

        return _currentTab.SelectedItems.Select(item => item.Entry.FullPath).ToList();
    }

    private bool CanOpenWithRightArrow(FileItemViewModel item)
    {
        if (item.Entry.IsDirectory)
            return true;

        if (!string.Equals(item.Entry.Extension, ".lnk", StringComparison.OrdinalIgnoreCase))
            return false;

        if (App.Services.GetRequiredService<Core.Interfaces.IShellIntegrationService>()
            .ResolveShortcutTarget(item.Entry.FullPath) is not string shortcutTarget)
        {
            return false;
        }

        return Directory.Exists(shortcutTarget);
    }

    private void ClearTransientSelection()
    {
        _typeAheadBuffer = string.Empty;
        _selectionAnchorIndex = -1;
        _isShiftPressed = false;

        if (_currentTab is not null)
        {
            _currentTab.SelectedItems.Clear();
        }

        _isClearingSelection = true;
        try
        {
            FileListView_Inner.SelectedItems.Clear();
            FileListView_Inner.SelectedItem = null;
        }
        finally
        {
            _isClearingSelection = false;
        }

        ApplySelectionFrames();
    }

    private bool TryApplyPendingSelectionAfterGoUp()
    {
        if (_currentTab is null || string.IsNullOrWhiteSpace(_pendingSelectPathAfterGoUp))
            return false;

        var pendingPath = NormalizePathForCompare(_pendingSelectPathAfterGoUp);
        var pendingName = Path.GetFileName(pendingPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));

        var targetItem = _currentTab.Items.FirstOrDefault(item =>
            string.Equals(NormalizePathForCompare(item.Entry.FullPath), pendingPath, StringComparison.OrdinalIgnoreCase));

        if (targetItem is null && !string.IsNullOrWhiteSpace(pendingName))
        {
            targetItem = _currentTab.Items.FirstOrDefault(item =>
                item.Entry.IsDirectory && string.Equals(item.Entry.Name, pendingName, StringComparison.OrdinalIgnoreCase));
        }

        if (targetItem is null)
            return false;

        _pendingSelectPathAfterGoUp = null;

        _isProgrammaticRangeSelection = true;
        try
        {
            FileListView_Inner.SelectedItems.Clear();
            FileListView_Inner.SelectedItem = targetItem;
        }
        finally
        {
            _isProgrammaticRangeSelection = false;
        }

        _selectionAnchorIndex = _currentTab.Items.IndexOf(targetItem);
        FileListView_Inner.ScrollIntoView(targetItem, ScrollIntoViewAlignment.Leading);
        FileListView_Inner.Focus(FocusState.Programmatic);
        ApplySelectionFrames();
        return true;
    }

    private void EnsureTopRowVisibleWhenNoSelection()
    {
        if (_currentTab is null || _currentTab.Items.Count == 0)
            return;

        if (!string.IsNullOrWhiteSpace(_pendingSelectPathAfterGoUp))
            return;

        if (FileListView_Inner.SelectedItem is not null)
            return;

        var first = _currentTab.Items[0];
        FileListView_Inner.ScrollIntoView(first, ScrollIntoViewAlignment.Leading);
    }

    private async void EnsurePendingSelectionAfterGoUpAsync()
    {
        if (string.IsNullOrWhiteSpace(_pendingSelectPathAfterGoUp))
            return;

        for (int attempt = 0; attempt < 12; attempt++)
        {
            await Task.Delay(35);
            if (TryApplyPendingSelectionAfterGoUp())
                return;
        }
    }

    private static string NormalizePathForCompare(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        var normalized = path;
        if (normalized.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            normalized = @"\\" + normalized[8..];
        else if (normalized.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            normalized = normalized[4..];

        try
        {
            normalized = Path.GetFullPath(normalized);
        }
        catch
        {
            // Keep original path if full-path normalization fails.
        }

        return normalized.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    #endregion

    #region Drag and Drop

    private void FileList_DragItemsStarting(object sender, DragItemsStartingEventArgs e)
    {
        var paths = e.Items
            .OfType<FileItemViewModel>()
            .Select(i => i.Entry.FullPath)
            .ToList();

        if (paths.Count == 0) return;

        var items = new List<IStorageItem>();
        foreach (var path in paths)
        {
            try
            {
                if (Directory.Exists(path))
                    items.Add(StorageFolder.GetFolderFromPathAsync(path).AsTask().Result);
                else if (File.Exists(path))
                    items.Add(StorageFile.GetFileFromPathAsync(path).AsTask().Result);
            }
            catch { /* Skip items that can't be accessed */ }
        }

        if (items.Count > 0)
        {
            e.Data.SetStorageItems(items);
            e.Data.RequestedOperation = DataPackageOperation.Copy | DataPackageOperation.Move;
        }
    }

    private void FileList_DragOver(object sender, DragEventArgs e)
    {
        if (e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            e.AcceptedOperation = e.Modifiers.HasFlag(Windows.ApplicationModel.DataTransfer.DragDrop.DragDropModifiers.Control)
                ? DataPackageOperation.Copy
                : DataPackageOperation.Move;
        }
    }

    private async void FileList_Drop(object sender, DragEventArgs e)
    {
        if (_currentTab is null) return;

        if (e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            var items = await e.DataView.GetStorageItemsAsync();
            var sourcePaths = items.Select(i => i.Path).ToList();
            var isMove = e.AcceptedOperation == DataPackageOperation.Move;

            var operation = new Core.Models.FileOperation
            {
                OperationType = isMove ? Core.Models.FileOperationType.Move : Core.Models.FileOperationType.Copy,
                SourcePaths = sourcePaths,
                DestinationPath = _currentTab.CurrentPath
            };

            var fileService = App.Services.GetRequiredService<Core.Interfaces.IFileSystemService>();
            await fileService.EnqueueOperationAsync(operation);
            await _currentTab.RefreshAsync();
        }
    }

    #endregion

    #region Empty State

    private void UpdateEmptyState()
    {
        bool isEmpty = _currentTab?.Items.Count == 0 && _currentTab?.IsLoading == false;
        EmptyStatePanel.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
    }

    #endregion

    #region Type-Ahead Search

    private void FileListView_Inner_CharacterReceived(UIElement sender, Microsoft.UI.Xaml.Input.CharacterReceivedRoutedEventArgs args)
    {
        if (_currentTab is null) return;

        if (_currentTab.Items.Any(item => item.IsRenaming))
        {
            _typeAheadBuffer = string.Empty;
            args.Handled = true;
            return;
        }

        var ch = args.Character;

        // Ignore control characters
        if (char.IsControl(ch)) return;

        var now = DateTimeOffset.UtcNow;
        var elapsed = (now - _lastKeyTime).TotalMilliseconds;

        // Reset buffer if more than 800ms since last keystroke
        if (elapsed > 800)
        {
            _typeAheadBuffer = string.Empty;
        }

        _lastKeyTime = now;
        var newBuffer = _typeAheadBuffer + ch;

        // Detect repeated same character (e.g. "sss") — cycle through matches
        var allSameChar = newBuffer.Length > 1 && newBuffer.All(c => char.ToLowerInvariant(c) == char.ToLowerInvariant(newBuffer[0]));

        if (allSameChar)
        {
            // Cycle: find all items starting with this char, pick the next one after current selection
            var prefix = newBuffer[0].ToString();
            var matches = _currentTab.Items
                .Where(item => item.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (matches.Count > 0)
            {
                var currentIndex = FileListView_Inner.SelectedItem is FileItemViewModel sel
                    ? matches.IndexOf(sel) : -1;
                var nextIndex = (currentIndex + 1) % matches.Count;
                var match = matches[nextIndex];

                var matchIndexInAllItems = _currentTab.Items.IndexOf(match);
                if (IsShiftSelectionActive() && matchIndexInAllItems >= 0)
                {
                    if (_selectionAnchorIndex < 0)
                        _selectionAnchorIndex = FileListView_Inner.SelectedItem is FileItemViewModel selected
                            ? _currentTab.Items.IndexOf(selected)
                            : matchIndexInAllItems;

                    ApplyRangeSelection(_selectionAnchorIndex, matchIndexInAllItems);
                }
                else
                {
                    _selectionAnchorIndex = matchIndexInAllItems;
                    FileListView_Inner.SelectedItem = match;
                }

                FileListView_Inner.ScrollIntoView(match, ScrollIntoViewAlignment.Leading);
            }

            _typeAheadBuffer = newBuffer;
        }
        else
        {
            _typeAheadBuffer = newBuffer;

            // Find first matching item for the accumulated buffer
            var match = _currentTab.Items
                .FirstOrDefault(item => item.Name.StartsWith(_typeAheadBuffer, StringComparison.OrdinalIgnoreCase));

            // If multi-char buffer didn't match, try single char
            if (match is null && _typeAheadBuffer.Length > 1)
            {
                var single = ch.ToString();
                match = _currentTab.Items
                    .FirstOrDefault(item => item.Name.StartsWith(single, StringComparison.OrdinalIgnoreCase));
                if (match is not null)
                    _typeAheadBuffer = single;
            }

            if (match is not null)
            {
                var matchIndexInAllItems = _currentTab.Items.IndexOf(match);
                if (IsShiftSelectionActive() && matchIndexInAllItems >= 0)
                {
                    if (_selectionAnchorIndex < 0)
                        _selectionAnchorIndex = FileListView_Inner.SelectedItem is FileItemViewModel selected
                            ? _currentTab.Items.IndexOf(selected)
                            : matchIndexInAllItems;

                    ApplyRangeSelection(_selectionAnchorIndex, matchIndexInAllItems);
                }
                else
                {
                    _selectionAnchorIndex = matchIndexInAllItems;
                    FileListView_Inner.SelectedItem = match;
                }

                FileListView_Inner.ScrollIntoView(match, ScrollIntoViewAlignment.Leading);
            }
        }

        args.Handled = true;
    }

    private FileItemViewModel? ResolveFolderTarget(FileItemViewModel? tappedItem)
    {
        if (_currentTab is null)
            return null;

        if (_currentTab.SelectedItems.Count == 1 && _currentTab.SelectedItems[0].Entry.IsDirectory)
            return _currentTab.SelectedItems[0];

        return tappedItem?.Entry.IsDirectory == true ? tappedItem : null;
    }

    #endregion

    #region Column Resize

    private void ColGrip_PointerEntered(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Border border)
        {
            border.Background = new SolidColorBrush(Windows.UI.Color.FromArgb(80, 128, 128, 128));
            ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast);
        }
    }

    private void ColGrip_PointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Border border)
        {
            border.Background = new SolidColorBrush(Colors.Transparent);
            ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.Arrow);
        }
    }

    private void ColGrip_Name_ManipulationDelta(object sender, ManipulationDeltaRoutedEventArgs e)
    {
        // Name column is star-sized; convert ActualWidth to pixel on first drag
        if (ColName.Width.IsStar)
            ColName.Width = new GridLength(ColName.ActualWidth);
        ResizeColumn(ColName, e.Delta.Translation.X, 100);
    }

    private void ColGrip_OneDrive_ManipulationDelta(object sender, ManipulationDeltaRoutedEventArgs e)
        => ResizeColumn(ColOneDrive, e.Delta.Translation.X, 30);

    private void ColGrip_DateModified_ManipulationDelta(object sender, ManipulationDeltaRoutedEventArgs e)
        => ResizeColumn(ColDateModified, e.Delta.Translation.X, 80);

    private void ColGrip_Type_ManipulationDelta(object sender, ManipulationDeltaRoutedEventArgs e)
        => ResizeColumn(ColType, e.Delta.Translation.X, 60);

    private void ResizeColumn(ColumnDefinition col, double delta, double minWidth)
    {
        var newWidth = col.Width.Value + delta;
        if (newWidth >= minWidth)
        {
            col.Width = new GridLength(newWidth);
            SyncRowColumnWidths();
        }
    }

    private void SyncRowColumnWidths()
    {
        // Walk all realized containers and update their inner Grid column widths
        for (int i = 0; i < FileListView_Inner.Items.Count; i++)
        {
            var container = FileListView_Inner.ContainerFromIndex(i) as ListViewItem;
            if (container is null) continue;
            ApplyColumnWidths(container);
            ApplySelectionFrame(container);
        }
    }

    private void FileListView_Inner_ContainerContentChanging(ListViewBase sender, ContainerContentChangingEventArgs args)
    {
        if (args.InRecycleQueue) return;
        // Register a callback for phase 1 when content is loaded
        args.RegisterUpdateCallback(OnContainerPhase1);
    }

    private void OnContainerPhase1(ListViewBase sender, ContainerContentChangingEventArgs args)
    {
        if (args.ItemContainer is not ListViewItem item) return;
        ApplyColumnWidths(item);
        ApplySelectionFrame(item);
    }

    private void ApplySelectionFrames()
    {
        for (int i = 0; i < FileListView_Inner.Items.Count; i++)
        {
            if (FileListView_Inner.ContainerFromIndex(i) is ListViewItem container)
            {
                ApplySelectionFrame(container);
            }
        }
    }

    private static void ApplySelectionFrame(ListViewItem container)
    {
        var border = FindSelectionFrameBorder(container);
        if (border is null)
            return;

        border.BorderThickness = container.IsSelected ? new Thickness(1) : new Thickness(0);
    }

    private static Border? FindSelectionFrameBorder(DependencyObject parent)
    {
        var count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is Border border &&
                border.Tag is string tag &&
                string.Equals(tag, "SelectionFrameBorder", StringComparison.Ordinal))
            {
                return border;
            }

            var found = FindSelectionFrameBorder(child);
            if (found is not null)
                return found;
        }

        return null;
    }

    private void ApplyColumnWidths(ListViewItem container)
    {
        var grid = FindDataRowGrid(container);
        if (grid is null || grid.ColumnDefinitions.Count < 5) return;

        // Mirror the header column widths exactly
        grid.ColumnDefinitions[0].Width = ColName.Width;
        grid.ColumnDefinitions[1].Width = new GridLength(ColOneDrive.Width.Value);
        grid.ColumnDefinitions[2].Width = new GridLength(ColDateModified.Width.Value);
        grid.ColumnDefinitions[3].Width = new GridLength(ColType.Width.Value);
        grid.ColumnDefinitions[4].Width = new GridLength(ColSize.Width.Value);
    }

    private static Grid? FindDataRowGrid(DependencyObject parent)
    {
        var count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is Grid g && g.ColumnDefinitions.Count == 5)
                return g;
            var found = FindDataRowGrid(child);
            if (found is not null) return found;
        }
        return null;
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

    private static T? FindAncestorOfType<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child is not null)
        {
            if (child is T match)
                return match;

            child = VisualTreeHelper.GetParent(child);
        }

        return null;
    }

    #endregion
}
