// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Core.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.ApplicationModel.DataTransfer;

namespace FileExplorer.App.Views;

/// <summary>
/// Right-side inspector containing identity fields, metadata, navigation hub trigger, and dump staging area.
/// </summary>
public sealed partial class InspectorSidebarView : UserControl
{
    private const string DumpFolderPath = @"C:\temp\x_file_dump";

    private sealed class NavigationHubRootListItem
    {
        public NavigationHubRootListItem(MainWindow.ShortcutHubRootItem root)
        {
            Root = root;
            DisplayName = root.DisplayName;
        }

        public MainWindow.ShortcutHubRootItem Root { get; }

        public string DisplayName { get; }
    }

    private TabContentViewModel? _boundTab;
    private FileItemViewModel? _selectedItem;
    private CancellationTokenSource? _folderAggregateCts;
    private readonly ObservableCollection<NavigationHubRootListItem> _navigationHubRoots = [];

    /// <summary>
    /// Initializes a new instance of the <see cref="InspectorSidebarView"/> class.
    /// </summary>
    public InspectorSidebarView()
    {
        InitializeComponent();
        NavigationHubRootList.ItemsSource = _navigationHubRoots;
        EnsureDumpFolderExists();
        ReloadNavigationHubRoots();
        _ = UpdateDumpSummaryAsync();
    }

    /// <summary>
    /// Binds this inspector to the currently active tab.
    /// </summary>
    public void BindToTab(TabContentViewModel tab)
    {
        if (_boundTab is not null)
        {
            _boundTab.SelectedItems.CollectionChanged -= SelectedItems_CollectionChanged;
            _boundTab.PropertyChanged -= BoundTab_PropertyChanged;
        }

        DetachSelectedItem();

        _boundTab = tab;
        _boundTab.SelectedItems.CollectionChanged += SelectedItems_CollectionChanged;
        _boundTab.PropertyChanged += BoundTab_PropertyChanged;

        UpdateSelectionBinding();
        ReloadNavigationHubRoots();
    }

    private void ReloadNavigationHubRoots()
    {
        _navigationHubRoots.Clear();

        if (App.MainWindow is not MainWindow mainWindow)
        {
            NavigationHubStatusText.Text = "Favorites hub unavailable.";
            SetDeepIndexWarning("rq_search engine unavailable: favorites hub missing.");
            return;
        }

        var roots = mainWindow.GetShortcutHubRootItems();
        foreach (var root in roots)
        {
            _navigationHubRoots.Add(new NavigationHubRootListItem(root));
        }

        NavigationHubStatusText.Text = roots.Count == 0
            ? "No favorites found."
            : "First-level favorites. Click any item for deeper levels.";

        _ = RefreshDeepIndexingStateAsync(roots);
    }

    private void BoundTab_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TabContentViewModel.CurrentPath) && _selectedItem is null)
        {
            UpdateMetadataView();
        }
    }

    private void SelectedItems_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        UpdateSelectionBinding();
    }

    private void UpdateSelectionBinding()
    {
        if (_boundTab is null)
            return;

        var selection = _boundTab.SelectedItems.ToList();
        if (selection.Count <= 1)
        {
            IdentitySinglePanel.Visibility = Visibility.Visible;
            IdentityAggregatePanel.Visibility = Visibility.Collapsed;
            AttachSelectedItem(selection.FirstOrDefault());
            return;
        }

        DetachSelectedItem();
        IdentitySinglePanel.Visibility = Visibility.Collapsed;
        IdentityAggregatePanel.Visibility = Visibility.Visible;

        UpdateAggregateSelectionView(selection);

        DateModifiedValue.Text = "Date: multiple selection";
        MetadataValue.Text = $"Details: {selection.Count} selected item(s)";
        PathValue.Text = $"Path: {DisplayOrDash(_boundTab.CurrentPath)}";
    }

    private void UpdateAggregateSelectionView(IReadOnlyList<FileItemViewModel> selection)
    {
        var files = selection.Where(item => !item.Entry.IsDirectory).ToList();
        var totalSize = files.Sum(item => item.Entry.Size);

        var typeSummary = selection
            .GroupBy(
                item => item.Entry.IsDirectory
                    ? "Folder"
                    : (string.IsNullOrWhiteSpace(item.Entry.Extension)
                        ? "(none)"
                        : item.Entry.Extension.ToLowerInvariant()),
                StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .Take(10)
            .Select(group => $"{group.Key}: {group.Count()}");

        AggregateSelectedCountValue.Text = $"Selected items: {selection.Count}";
        AggregateTotalSizeValue.Text = $"Total size: {FormatBytes(totalSize)}";
        AggregateByTypeValue.Text = $"Types: {string.Join(", ", typeSummary)}";
    }

    private void AttachSelectedItem(FileItemViewModel? item)
    {
        DetachSelectedItem();
        _selectedItem = item;

        if (_selectedItem is not null)
        {
            _selectedItem.PropertyChanged += SelectedItem_PropertyChanged;
        }

        UpdateIdentityView();
        UpdateMetadataView();

        if (_selectedItem?.Entry.IsDirectory == true)
        {
            _ = LoadFolderAggregateAsync(_selectedItem.Entry.FullPath);
        }
    }

    private void DetachSelectedItem()
    {
        _folderAggregateCts?.Cancel();
        _folderAggregateCts?.Dispose();
        _folderAggregateCts = null;

        if (_selectedItem is not null)
        {
            _selectedItem.PropertyChanged -= SelectedItem_PropertyChanged;
        }

        _selectedItem = null;
    }

    private void SelectedItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (IdentityAggregatePanel.Visibility == Visibility.Visible)
            return;

        if (e.PropertyName is nameof(FileItemViewModel.ExtractedProject) or
            nameof(FileItemViewModel.ExtractedRevision) or
            nameof(FileItemViewModel.ExtractedDocumentName))
        {
            UpdateIdentityView();
            return;
        }

        if (e.PropertyName is nameof(FileItemViewModel.DateModifiedText) or
            nameof(FileItemViewModel.SizeText) or
            nameof(FileItemViewModel.FullPath))
        {
            UpdateMetadataView();

            if (_selectedItem?.Entry.IsDirectory == true)
            {
                _ = LoadFolderAggregateAsync(_selectedItem.Entry.FullPath);
            }
        }
    }

    private void UpdateIdentityView()
    {
        ProjectValue.Text = DisplayOrDash(_selectedItem?.ExtractedProject);
        RevisionValue.Text = DisplayOrDash(_selectedItem?.ExtractedRevision);
        DocumentNameValue.Text = DisplayOrDash(_selectedItem?.ExtractedDocumentName);
    }

    private void UpdateMetadataView()
    {
        if (_selectedItem is null)
        {
            DateModifiedValue.Text = "Date: -";
            MetadataValue.Text = "Details: -";
            PathValue.Text = $"Path: {DisplayOrDash(_boundTab?.CurrentPath)}";
            return;
        }

        DateModifiedValue.Text = $"Date: {_selectedItem.Entry.DateModified.LocalDateTime:dd.MM.yyyy HH:mm}";

        if (_selectedItem.Entry.IsDirectory)
        {
            MetadataValue.Text = "Details: Calculating recursive size and file count...";
        }
        else
        {
            MetadataValue.Text = $"Details: {_selectedItem.Entry.Extension} | {FormatBytes(_selectedItem.Entry.Size)}";
        }

        PathValue.Text = $"Path: {_selectedItem.Entry.FullPath}";
    }

    private async Task LoadFolderAggregateAsync(string folderPath)
    {
        _folderAggregateCts?.Cancel();
        _folderAggregateCts?.Dispose();
        _folderAggregateCts = new CancellationTokenSource();
        var ct = _folderAggregateCts.Token;

        MetadataValue.Text = "Details: Calculating recursive size and file count...";

        try
        {
            var rqUtility = App.Services.GetRequiredService<IRqUtilityService>();
            var aggregate = await rqUtility.GetFolderAggregateAsync(folderPath, ct);

            if (_selectedItem?.Entry.FullPath != folderPath)
                return;

            if (!aggregate.IsSuccess)
            {
                MetadataValue.Text = $"Details: Folder | aggregation failed ({aggregate.ErrorMessage})";
                return;
            }

            var source = aggregate.UsedRq ? "rq" : "fallback";
            MetadataValue.Text =
                $"Details: Folder | Files: {aggregate.FileCount} | Size: {FormatBytes(aggregate.TotalSizeBytes)} | {source} {aggregate.ElapsedMilliseconds} ms";
        }
        catch (OperationCanceledException)
        {
            // Superseded by a newer selection.
        }
        catch (Exception ex)
        {
            if (_selectedItem?.Entry.FullPath == folderPath)
            {
                MetadataValue.Text = $"Details: Folder | aggregation failed ({ex.Message})";
            }
        }
    }

    private async void DumpDropBorder_Drop(object sender, DragEventArgs e)
    {
        if (!e.DataView.Contains(StandardDataFormats.StorageItems))
            return;

        var storageItems = await e.DataView.GetStorageItemsAsync();
        var sourcePaths = storageItems
            .Select(item => item.Path)
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .ToList();

        if (sourcePaths.Count == 0)
            return;

        EnsureDumpFolderExists();

        var operation = new FileOperation
        {
            OperationType = FileOperationType.Move,
            SourcePaths = sourcePaths,
            DestinationPath = DumpFolderPath
        };

        var fileService = App.Services.GetRequiredService<IFileSystemService>();

        var stopwatch = Stopwatch.StartNew();
        await fileService.EnqueueOperationAsync(operation);
        stopwatch.Stop();

        Debug.WriteLine($"Dump staging move completed in {stopwatch.ElapsedMilliseconds} ms for {sourcePaths.Count} item(s).");

        if (_boundTab is not null)
        {
            await _boundTab.RefreshAsync();
        }

        await UpdateDumpSummaryAsync();
    }

    private void DumpDropBorder_DragOver(object sender, DragEventArgs e)
    {
        if (e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            e.AcceptedOperation = DataPackageOperation.Move;
        }
    }

    private async Task UpdateDumpSummaryAsync()
    {
        EnsureDumpFolderExists();

        var stopwatch = Stopwatch.StartNew();
        var (count, sizeBytes) = await Task.Run(() =>
        {
            var directory = new DirectoryInfo(DumpFolderPath);
            var files = directory.EnumerateFiles("*", SearchOption.AllDirectories).ToList();
            return (files.Count, files.Sum(file => file.Length));
        });
        stopwatch.Stop();

        DumpSummaryText.Text = $"Items: {count} | Size: {FormatBytes(sizeBytes)} | elapsed {stopwatch.ElapsedMilliseconds} ms";
    }

    private void NavigationHubRootList_ItemClick(object sender, ItemClickEventArgs e)
    {
        if (App.MainWindow is not MainWindow mainWindow)
            return;

        if (e.ClickedItem is not NavigationHubRootListItem clicked)
            return;

        if (sender is ListView listView && listView.ContainerFromItem(clicked) is FrameworkElement anchor)
        {
            mainWindow.ShowShortcutHubChildrenFlyout(clicked.Root, anchor);
            return;
        }

        mainWindow.ShowShortcutHubChildrenFlyout(clicked.Root, NavigationHubRootList);
    }

    private static string FormatBytes(long bytes)
    {
        return bytes switch
        {
            < 1024 => $"{bytes} B",
            < 1024 * 1024 => $"{bytes / 1024.0:F1} KB",
            < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):F1} MB",
            _ => $"{bytes / (1024.0 * 1024 * 1024):F2} GB"
        };
    }

    private static string DisplayOrDash(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
    }

    private async Task RefreshDeepIndexingStateAsync(IReadOnlyList<MainWindow.ShortcutHubRootItem> roots)
    {
        try
        {
            var deepIndexingService = App.Services.GetRequiredService<IDeepIndexingService>();
            var rootPaths = roots
                .Where(root => root.IsFolder && !string.IsNullOrWhiteSpace(root.FolderPath))
                .Select(root => root.FolderPath!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            await deepIndexingService.StartOrRefreshAsync(rootPaths);
            SetDeepIndexWarning(deepIndexingService.EngineWarning);
        }
        catch (Exception ex)
        {
            SetDeepIndexWarning(ex.Message);
        }
    }

    private void SetDeepIndexWarning(string? warning)
    {
        if (string.IsNullOrWhiteSpace(warning))
        {
            DeepIndexWarningText.Visibility = Visibility.Collapsed;
            DeepIndexWarningText.Text = string.Empty;
            return;
        }

        DeepIndexWarningText.Text = $"Deep indexing warning: {warning}";
        DeepIndexWarningText.Visibility = Visibility.Visible;
    }

    private static void EnsureDumpFolderExists()
    {
        if (!Directory.Exists(DumpFolderPath))
        {
            Directory.CreateDirectory(DumpFolderPath);
        }
    }
}
