// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using System.IO;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Core.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;

namespace FileExplorer.App.Views;

/// <summary>
/// Right-side navigation hub containing shortcut roots that open flyout menus.
/// </summary>
public sealed partial class InspectorSidebarView : UserControl
{
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

    private sealed class FavoriteListItem
    {
        public FavoriteListItem(FavoriteEntrySettings favorite)
        {
            FullPath = favorite.Path;
            DefaultName = GetPathLeafLabel(favorite.Path);
            Alias = favorite.Alias;
            IsDirectory = Directory.Exists(favorite.Path);
            IconGlyph = IsDirectory ? "\uE8B7" : "\uE8A5";
        }

        public string FullPath { get; }

        public string DefaultName { get; }

        public string? Alias { get; }

        public string DisplayName => string.IsNullOrWhiteSpace(Alias) ? DefaultName : Alias;

        public bool IsDirectory { get; }

        public string IconGlyph { get; }

        private static string GetPathLeafLabel(string path)
        {
            var trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var leaf = Path.GetFileName(trimmed);
            return string.IsNullOrWhiteSpace(leaf) ? path : leaf;
        }
    }

    private readonly ObservableCollection<NavigationHubRootListItem> _navigationHubRoots = [];
    private readonly ObservableCollection<FavoriteListItem> _favorites = [];
    private readonly ISettingsService _settingsService;
    private TabContentViewModel? _currentTab;

    /// <summary>
    /// Initializes a new instance of the <see cref="InspectorSidebarView"/> class.
    /// </summary>
    public InspectorSidebarView()
    {
        InitializeComponent();
        _settingsService = App.Services.GetRequiredService<ISettingsService>();
        NavigationHubRootList.ItemsSource = _navigationHubRoots;
        FavoritesList.ItemsSource = _favorites;
        ReloadNavigationHubRoots();
        ReloadFavorites();
    }

    /// <summary>
    /// Binds this inspector to the currently active tab.
    /// </summary>
    public void BindToTab(TabContentViewModel tab)
    {
        if (ReferenceEquals(_currentTab, tab))
            return;

        _currentTab = tab;
        ReloadNavigationHubRoots();
        ReloadFavorites();
    }

    public void RefreshFavorites()
    {
        ReloadFavorites();
    }

    private void ReloadNavigationHubRoots()
    {
        _navigationHubRoots.Clear();

        if (App.MainWindow is not MainWindow mainWindow)
            return;

        var roots = mainWindow.GetShortcutHubRootItems();
        foreach (var root in roots)
        {
            _navigationHubRoots.Add(new NavigationHubRootListItem(root));
        }
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

    private void ReloadFavorites()
    {
        _favorites.Clear();

        foreach (var favorite in _settingsService.Settings.Favorites)
        {
            _favorites.Add(new FavoriteListItem(favorite));
        }
    }

    private void FavoritesList_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
    {
        var favorite = ResolveFavoriteFromOriginalSource(e.OriginalSource as DependencyObject)
            ?? FavoritesList.SelectedItem as FavoriteListItem;

        if (favorite is null)
            return;

        if (App.MainWindow is not MainWindow mainWindow)
            return;

        _ = mainWindow.OpenFavoritePathAsync(favorite.FullPath);
    }

    private void FavoritesList_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (App.MainWindow is not MainWindow mainWindow)
            return;

        var clickedFavorite = ResolveFavoriteFromOriginalSource(e.OriginalSource as DependencyObject);
        if (clickedFavorite is not null && !FavoritesList.SelectedItems.Contains(clickedFavorite))
        {
            FavoritesList.SelectedItems.Clear();
            FavoritesList.SelectedItem = clickedFavorite;
        }

        var selectedFavorites = FavoritesList.SelectedItems.OfType<FavoriteListItem>().ToList();
        if (selectedFavorites.Count == 0 && clickedFavorite is not null)
        {
            selectedFavorites.Add(clickedFavorite);
        }

        if (selectedFavorites.Count == 0)
            return;

        var flyout = new MenuFlyout();

        if (selectedFavorites.Count == 1)
        {
            var renameItem = new MenuFlyoutItem
            {
                Text = "Rename favorite",
                Icon = new FontIcon { Glyph = "\uE8AC" }
            };
            renameItem.Click += async (_, _) =>
            {
                var alias = await PromptFavoriteAliasAsync(selectedFavorites[0]);
                if (alias is null)
                    return;

                await mainWindow.RenameFavoriteAsync(selectedFavorites[0].FullPath, alias);
            };
            flyout.Items.Add(renameItem);
            flyout.Items.Add(new MenuFlyoutSeparator());
        }

        var removeText = selectedFavorites.Count == 1
            ? "Remove favorite"
            : $"Remove favorites ({selectedFavorites.Count})";

        var removeItem = new MenuFlyoutItem
        {
            Text = removeText,
            Icon = new FontIcon { Glyph = "\uE74D" }
        };
        removeItem.Click += async (_, _) =>
        {
            var paths = selectedFavorites.Select(item => item.FullPath).ToList();
            await mainWindow.RemoveFavoritesAsync(paths);
        };
        flyout.Items.Add(removeItem);

        flyout.ShowAt(FavoritesList, e.GetPosition(FavoritesList));
    }

    private async Task<string?> PromptFavoriteAliasAsync(FavoriteListItem favorite)
    {
        var editor = new TextBox
        {
            Text = favorite.DisplayName,
            PlaceholderText = "Favorite name",
            MinWidth = 320
        };

        var dialog = new ContentDialog
        {
            XamlRoot = XamlRoot,
            Title = "Rename favorite",
            Content = editor,
            PrimaryButtonText = "Save",
            SecondaryButtonText = "Use original name",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary
        };

        var result = await dialog.ShowAsync();
        return result switch
        {
            ContentDialogResult.Primary => editor.Text,
            ContentDialogResult.Secondary => string.Empty,
            _ => null
        };
    }

    private static FavoriteListItem? ResolveFavoriteFromOriginalSource(DependencyObject? source)
    {
        while (source is not null)
        {
            if (source is FrameworkElement element && element.DataContext is FavoriteListItem favorite)
                return favorite;

            source = VisualTreeHelper.GetParent(source);
        }

        return null;
    }
}
