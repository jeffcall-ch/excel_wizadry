# FileExplorer – Agent Handoff Brief

## What This Is

A Windows 11 File Explorer replacement built with WinUI 3, .NET 8, C# 12, CommunityToolkit.Mvvm. All source code has been written (56 files across 4 projects). The project **builds successfully**.

## Solution Structure

```
FileExplorer/
├── FileExplorer.sln
├── Directory.Build.props          # LangVersion=12, Nullable, TreatWarningsAsErrors=true
├── FileExplorer.Core/             # ViewModels, Models, Services interfaces (BUILDS ✅)
├── FileExplorer.Infrastructure/   # Native interop, service implementations (BUILDS ✅)
├── FileExplorer.Tests/            # xUnit tests (BUILDS ✅)
└── FileExplorer.App/              # WinUI 3 app, XAML, DI setup (BUILDS ✅)
```

Build command: `dotnet build FileExplorer.sln -p:Platform=x64`

## What Works

- **All 4 projects** – compile for x64 with 0 warnings and 0 errors.

## XAML Compiler Workarounds Applied

### Problem: XamlCompiler.exe crashes when resolving WinUI types in App.xaml

The WinUI XAML compiler (`XamlCompiler.exe`, .NET Framework 4.7.2 process) crashes silently (exit code 1, no output.json) when XAML contains type references that require resolving WinRT projection assemblies — specifically `<XamlControlsResources xmlns="using:Microsoft.UI.Xaml.Controls" />` and `<Style TargetType="Button">` in App.xaml.

### Workarounds

1. **XamlControlsResources loaded in code-behind** — Moved from App.xaml to App.xaml.cs constructor: `Resources.MergedDictionaries.Add(new XamlControlsResources());`

2. **SubtleButtonStyle defined in code-behind** — The `<Style TargetType="Button">` in App.xaml also triggers the crash. Defined programmatically in App constructor.

3. **x:Bind with Converter changed to Binding** — `{x:Bind ViewModel.ShowPreviewPane, Converter=...}` in MainWindow.xaml generates `SetConverterLookupRoot(this)` which requires `FrameworkElement`. In WinUI 3, `Window` does not inherit from `FrameworkElement`, causing CS1503. Changed to `{Binding}`.

4. **PRI generation stubs** — No Visual Studio Build Tools installed. Stub task DLLs placed at `%DOTNET_SDK%\Microsoft\VisualStudio\v17.0\AppxPackage\` to satisfy `Microsoft.Build.Packaging.Pri.Tasks.dll` and `Microsoft.Build.AppxPackage.dll` references. These are no-op implementations.

5. **AppxGeneratePriEnabled=false** and **DisableXbfGeneration=true** — Skip PRI/XBF pipeline that requires VS Build Tools.

### Root Cause

The `.NET Framework 4.7.2` XamlCompiler.exe process crashes when loading WinRT type metadata from projection assemblies (WinRT.Runtime.dll etc.) during type resolution in App.xaml. This does NOT occur when type references are only in Page/UserControl XAML. The in-process XAML compiler (`CompileXaml` MSBuild task) also fails with "Operation is not supported on this platform" on .NET Core runtime.

## Fixes Already Applied

| Issue | Fix |
|---|---|
| NuGet version errors | Pinned to `WindowsAppSDK 1.6.250228001`, replaced `CommunityToolkit.WinUI.UI.Controls` with `CommunityToolkit.WinUI.Collections 8.2.251219` |
| SYSLIB1051 P/Invoke | Changed `LibraryImport` → `DllImport` for `MoveFileEx` and `GetFileAttributesEx` in `NativeMethods.cs` |
| IL2026/IL2050 trimming | Removed `<IsTrimmable>` from Infrastructure.csproj |
| CS0219 unused variable | Removed `unpinned` from `CloudStatusService.cs` |
| Protected OnPropertyChanged | Added public `NotifyPropertyChanged()` wrapper in `FileItemViewModel` |
| XAML: Flyout as Grid child | Moved to `FlyoutBase.AttachedFlyout` in `FileListView.xaml` |
| XAML: TreeViewItem in DataTemplate | Flattened to StackPanel |
| XAML: `Cursor="SizeWestEast"` | Removed (not valid in WinUI 3) |
| PRI generation error | Set `<WindowsPackageType>None</WindowsPackageType>` |
| Missing using | Added `using Windows.Foundation` to `MainWindow.xaml.cs` for `TypedEventHandler` |
| XAML compiler crash | Moved `XamlControlsResources` and `SubtleButtonStyle` from App.xaml to App.xaml.cs code-behind |
| x:Bind converter in Window | Changed `{x:Bind ... Converter}` to `{Binding}` in MainWindow.xaml (Window ≠ FrameworkElement) |
| PRI generation missing tasks | Added `AppxGeneratePriEnabled=false`, `DisableXbfGeneration=true`; installed stub task DLLs |
| FlyoutBase missing using | Added `using Microsoft.UI.Xaml.Controls.Primitives` to FileListView.xaml.cs |

## Removed Packages (not restored)

- `Microsoft.Extensions.Logging.Debug` – removed during diagnosis. The `builder.AddDebug()` call was also removed from `App.xaml.cs`. Re-add if desired.
- `CommunityToolkit.WinUI.Controls.Primitives` – was never actually used in code.

## Prerequisites for Building on a New Machine

The build requires stub MSBuild task DLLs at:
```
%DOTNET_SDK_PATH%\Microsoft\VisualStudio\v17.0\AppxPackage\Microsoft.Build.Packaging.Pri.Tasks.dll
%DOTNET_SDK_PATH%\Microsoft\VisualStudio\v17.0\AppxPackage\Microsoft.Build.AppxPackage.dll
```
These are no-op implementations of PRI and Appx tasks. On the current machine they are at:
`C:\Users\szil\AppData\Local\Microsoft\dotnet\sdk\8.0.411\Microsoft\VisualStudio\v17.0\AppxPackage\`

Alternatively, install **Visual Studio Build Tools** with the Windows App SDK workload, which provides the real task assemblies.

## Environment

- Windows 11, .NET 8 SDK, **no Visual Studio installed** (only VS Code)
- NuGet cache: `C:\Users\szil\.nuget\packages\`
- Python 3.x available
