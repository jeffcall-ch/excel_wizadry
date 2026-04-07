// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.ObjectModel;
using FileExplorer.Core.ViewModels;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

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

    private readonly ObservableCollection<NavigationHubRootListItem> _navigationHubRoots = [];
    private TabContentViewModel? _currentTab;

    /// <summary>
    /// Initializes a new instance of the <see cref="InspectorSidebarView"/> class.
    /// </summary>
    public InspectorSidebarView()
    {
        InitializeComponent();
        NavigationHubRootList.ItemsSource = _navigationHubRoots;
        ReloadNavigationHubRoots();
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
}
