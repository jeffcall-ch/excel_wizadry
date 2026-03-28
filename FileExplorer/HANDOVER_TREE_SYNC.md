# Handover: WinUI 3 TreeView Sync Bug

## Problem
When the user navigates to a folder via the **right pane** (file list — double-click a folder), the **left pane** (TreeView) does not update. It should expand the tree to show and highlight the current folder, but it stays collapsed/unselected.

Navigation from the **left pane → right pane works fine** (clicking a tree node updates the file list). The reverse direction is broken.

## Architecture Overview

### Tech Stack
- **WinUI 3 / Windows App SDK 1.6**, .NET 8, self-contained
- **DisableXbfGeneration=true** — DataTemplates don't compile to XBF
- **CommunityToolkit.Mvvm 8.3.2** — ObservableObject, RelayCommand
- **Build**: `dotnet build FileExplorer.App\FileExplorer.App.csproj -c Debug -p:Platform=x64`
- **EXE**: `FileExplorer.App\bin\x64\Debug\net8.0-windows10.0.22621.0\FileExplorer.exe`

### Key Files
| File | Role |
|------|------|
| `FileExplorer.App\MainWindow.xaml.cs` | WinUI TreeView control, event handlers, sync logic |
| `FileExplorer.App\MainWindow.xaml` | TreeView XAML (line ~206): `<TreeView x:Name="NavigationTree" ...>` |
| `FileExplorer.Core\ViewModels\TreeViewModel.cs` | Tree ViewModel: `SyncWithPathAsync()`, `ExpandAndSelectAsync()`, `SyncCompleted` event |
| `FileExplorer.Core\ViewModels\MainViewModel.cs` | Wires `tab.Navigated += OnTabNavigated` which calls `TreeView.SyncWithPath(path)` |
| `FileExplorer.Core\ViewModels\TabContentViewModel.cs` | Fires `Navigated` event after navigation completes |

### Data Flow (supposed to work)
```
User double-clicks folder in file list
  → TabContentViewModel.OpenItemAsync() → NavigateAsync() → NavigateCoreAsync()
    → Enumerates directory, populates Items
    → Fires: Navigated?.Invoke(this, path)  [line ~174]

MainViewModel.OnTabNavigated(sender, path)  [line ~498]
  → TreeView.SyncWithPath(path)  [fire-and-forget wrapper]
    → TreeViewModel.SyncWithPathAsync(path)
      → ExpandAndSelectAsync() walks tree nodes, loads children, sets IsExpanded/SelectedNode
      → Fires: SyncCompleted?.Invoke(this, SelectedNode)

MainWindow.OnTreeSyncCompleted(sender, selectedVm)
  → Dispatched to UI thread
  → PopulateTreeView() — rebuilds WinUI TreeViewNode objects from VM state
  → CollectExpandedNodes() — gathers nodes to expand (bottom-up/deepest first)
  → Sets IsExpanded = true on each
  → Deferred dispatch: finds matching TreeViewNode, sets NavigationTree.SelectedNode
```

## What Has Been Tried (and failed)

### Attempt 1: Set IsExpanded in CreateTreeViewNode constructor
```csharp
var node = new TreeViewNode { IsExpanded = viewModel.IsExpanded };
```
**Result**: WinUI ignores `IsExpanded` set on TreeViewNodes before they're added to the live tree. Nodes appear collapsed.

### Attempt 2: ApplyExpansionState (top-down walk after PopulateTreeView)
After `PopulateTreeView()`, walk all nodes and set `IsExpanded = true` from ViewModel.
**Result**: WinUI fires `Expanding` event for each node. The `NavigationTree_Expanding` handler is `async void` — it calls `ExpandNodeAsync()`, clears children, re-adds them with NEW objects. These new child objects don't match the ones `FindTreeViewNode` is looking for. Selection fails.

### Attempt 3: Incremental SyncTreeViewNodes (no full rebuild)
Instead of rebuilding, walk existing tree nodes and add missing ones, sync children.
**Result**: Complex and fragile. The `Expanding` event still fires when setting `IsExpanded`, causing the same child-replacement race. Also, dictionary lookup by ViewModel reference fails when nodes are recreated.

### Attempt 4: _isSyncingTree flag + full rebuild + bottom-up expansion (current)
- Set `_isSyncingTree = true` to suppress `NavigationTree_Expanding` handler
- `PopulateTreeView()` full rebuild
- `CollectExpandedNodes()` deepest-first, then set `IsExpanded = true`
- Deferred `NavigationTree.SelectedNode = tvNode`

**Result**: Still not working. Possible causes:
1. The `_isSyncingTree` flag may not actually suppress the `Expanding` event (it fires asynchronously from WinUI's layout system, possibly after we set the flag back to false)
2. `PopulateTreeView` creates all nodes with `HasUnrealizedChildren = true` for loaded+has-children nodes, but since we DON'T set `IsExpanded` in the constructor, WinUI might not show children even though they're added
3. The `SyncCompleted` event might not fire at all if `SelectedNode` is null (check: `if (SelectedNode is not null)` guard in `SyncWithPathAsync` line ~101)
4. The fire-and-forget `_ = SyncWithPathSafeAsync(path)` might swallow exceptions silently
5. `ExpandAndSelectAsync` might fail to match paths (normalization issue with drive root "C:" vs "C:\")

## Suspected Root Causes to Investigate

### 1. SyncCompleted never fires
`SyncWithPathAsync` only fires `SyncCompleted` if `SelectedNode is not null`. If `ExpandAndSelectAsync` fails to find the target path (e.g., path normalization mismatch), `SelectedNode` stays null and the event never fires. **Add logging/breakpoint here first.**

### 2. WinUI TreeView.IsExpanded timing
Even with `_isSyncingTree`, setting `IsExpanded = true` on a `TreeViewNode` might trigger the `Expanding` event **after** `_isSyncingTree` is set back to `false` (since WinUI processes property changes asynchronously during layout). The guard might not help.

### 3. HasUnrealizedChildren conflict
`CreateTreeViewNode` sets `HasUnrealizedChildren = viewModel.HasChildren && !viewModel.IsLoaded`. But after `SyncWithPathAsync`, the intermediate nodes ARE loaded (it called `ExpandNodeAsync` on them). So `HasUnrealizedChildren` should be `false` and children should be populated. **Verify this is actually the case by checking `IsLoaded` on the VMs after sync.**

### 4. Fire-and-forget async gap
`SyncWithPath` does `_ = SyncWithPathSafeAsync(path)`. This async task runs on a thread pool thread. The `SyncCompleted` event fires from that context. The `OnTreeSyncCompleted` handler uses `_dispatcherQueue.TryEnqueue()` to get back to the UI thread. **If the DispatcherQueue is busy or the event is lost, nothing happens.**

## Recommended Debugging Approach

1. **Add a temp debug log** at the start of `SyncWithPathAsync`, at the `SyncCompleted?.Invoke` line, and at the start of `OnTreeSyncCompleted`. Write to `System.Diagnostics.Debug.WriteLine()` or a file at `%TEMP%\fe_tree_debug.log`. This will tell you whether the chain even executes.

2. **Check if `ExpandAndSelectAsync` finds the target path**. Log the `targetPath` and each `normalizedNodePath` it compares. A drive like `C:\` normalizes to `C:` after `TrimEnd('\\')` — this could break the prefix check.

3. **Try NOT suppressing the Expanding event**. Instead, make the `Expanding` handler a no-op when children are already loaded:
```csharp
if (!node.IsLoaded) { /* load children */ }
// If already loaded, just let WinUI show existing children
```

4. **Try using `await` instead of fire-and-forget** for the sync. In `OnTabNavigated`, use `_ = SyncAndUpdateTreeAsync(path)` that awaits `SyncWithPathAsync` and then dispatches update to UI thread directly, rather than relying on the event.

## Key Code Locations

- **TreeView XAML**: `MainWindow.xaml` line 206-213
- **PopulateTreeView()**: `MainWindow.xaml.cs` line ~378
- **CreateTreeViewNode()**: `MainWindow.xaml.cs` line ~386
- **OnTreeSyncCompleted()**: `MainWindow.xaml.cs` line ~484
- **NavigationTree_Expanding**: `MainWindow.xaml.cs` line ~433
- **SyncWithPathAsync**: `TreeViewModel.cs` line ~87
- **ExpandAndSelectAsync**: `TreeViewModel.cs` line ~122
- **OnTabNavigated**: `MainViewModel.cs` line ~498
- **TreeNodeViewModel class**: `TreeViewModel.cs` line ~380
