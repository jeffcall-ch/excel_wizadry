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

    /// <summary>Initializes a new instance of the <see cref="FileListView"/> class.</summary>
    public FileListView()
    {
        InitializeComponent();
    }

    /// <summary>Binds this view to a tab's content ViewModel.</summary>
    public void BindToTab(TabContentViewModel tab)
    {
        _currentTab = tab;
        FileListView_Inner.ItemsSource = tab.Items;
        UpdateEmptyState();
        UpdateSortIndicators();
        UpdateOneDriveColumnVisibility();

        // Subscribe to collection changes for empty state
        tab.Items.CollectionChanged += (_, _) =>
        {
            DispatcherQueue.TryEnqueue(() =>
            {
                UpdateEmptyState();
            });
        };

        // Subscribe to loading state
        tab.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(TabContentViewModel.IsLoading))
            {
                DispatcherQueue.TryEnqueue(() =>
                {
                    LoadingRing.Visibility = tab.IsLoading ? Visibility.Visible : Visibility.Collapsed;
                    LoadingRing.IsActive = tab.IsLoading;
                });
            }
        };
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
            ColOneDriveCheck.IsChecked = false;
        }
    }

    #endregion

    #region Item Interactions

    private void FileList_ItemClick(object sender, ItemClickEventArgs e)
    {
        // Single click selects (handled by ListView default behavior)
    }

    private void FileList_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (_currentTab is null) return;
        var item = FileListView_Inner.SelectedItem as FileItemViewModel;
        if (item is not null)
        {
            _ = _currentTab.OpenItemAsync(item);
        }
    }

    private void FileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_currentTab is null) return;
        _currentTab.SelectedItems.Clear();
        foreach (var item in FileListView_Inner.SelectedItems.OfType<FileItemViewModel>())
        {
            _currentTab.SelectedItems.Add(item);
        }
    }

    /// <summary>Sets whether the list is showing search results.</summary>
    public void SetSearchMode(bool isSearch) => _isSearchMode = isSearch;

    private void FileList_RightTapped(object sender, RightTappedRoutedEventArgs e)
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

        var flyout = new MenuFlyout();

        // Standard items
        flyout.Items.Add(new MenuFlyoutItem { Text = "Open", Icon = new FontIcon { Glyph = "\uE8E5" } });
        ((MenuFlyoutItem)flyout.Items[^1]).Click += (_, _) =>
        {
            var sel = _currentTab.SelectedItems.FirstOrDefault();
            if (sel is not null) _ = _currentTab.OpenItemAsync(sel);
        };

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
            _ = _currentTab.DeleteSelectedAsync();
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
            if (App.MainWindow is MainWindow mw)
            {
                var mainVm = App.Services.GetRequiredService<MainViewModel>();
                mainVm.ShowProperties(mw.Hwnd);
            }
        };

        flyout.ShowAt(FileListView_Inner, e.GetPosition(FileListView_Inner));
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

    #region Row Hover

    private void Row_PointerEntered(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Grid grid)
        {
            // System accent color at 10% opacity
            var accentColor = (Windows.UI.Color)Application.Current.Resources["SystemAccentColor"];
            grid.Background = new SolidColorBrush(Windows.UI.Color.FromArgb(25, accentColor.R, accentColor.G, accentColor.B));
        }
    }

    private void Row_PointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (sender is Grid grid)
        {
            grid.Background = new SolidColorBrush(Colors.Transparent);
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
                FileListView_Inner.SelectedItem = match;
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
                FileListView_Inner.SelectedItem = match;
                FileListView_Inner.ScrollIntoView(match, ScrollIntoViewAlignment.Leading);
            }
        }

        args.Handled = true;
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

    #endregion
}
