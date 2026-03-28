# FileExplorer — Windows 11 File Explorer Replacement

A production-grade Windows 11 File Explorer replacement built with WinUI 3, .NET 8, and CommunityToolkit.Mvvm.

## Architecture

```
FileExplorer.sln
├── FileExplorer.App             # WinUI 3 entry point, DI root, XAML views
├── FileExplorer.Core            # Domain models, interfaces, ViewModels (no UI refs)
├── FileExplorer.Infrastructure  # Win32/COM P/Invoke, Cloud Filter API, services
└── FileExplorer.Tests           # xUnit unit tests
```

## Build

### Prerequisites

- Visual Studio 2022 17.9+ or .NET 8 SDK
- Windows App SDK 1.5+ workload
- Windows 11 SDK (10.0.22621.0)

### Build from command line

```powershell
cd FileExplorer
dotnet restore
dotnet build -c Release -p:Platform=x64
```

### Run tests

```powershell
dotnet test FileExplorer.Tests
```

### Publish — MSIX

```powershell
dotnet publish FileExplorer.App -c Release -p:Platform=x64 /p:PublishProfile=Properties\PublishProfiles\MSIX-x64.pubxml
```

### Publish — Native AOT (single file)

```powershell
dotnet publish FileExplorer.App -c Release -p:Platform=x64 /p:PublishProfile=Properties\PublishProfiles\AOT-x64.pubxml
```

## Key Features (v1)

- **Tabbed browsing** with per-tab navigation history, FileSystemWatcher, and sort state
- **Native shell context menus** via IContextMenu COM interop (includes 7-Zip, Defender, etc.)
- **OneDrive integration** via Cloud Filter API with non-blocking status resolution
- **Channel-based operation queue** for async Copy/Move/Delete with progress and conflict resolution
- **Optimistic UI** with rollback on rename, delete, and new folder creation
- **Undo stack** for reversible file operations (Ctrl+Z)
- **FileSystemWatcher** with 250ms debounced change coalescing
- **Preview pane** for images, text, PDF, and shell thumbnails
- **Full keyboard shortcut parity** with Windows Explorer
- **Mica backdrop** with automatic Light/Dark theme support
- **Session persistence** — tabs and settings restore on next launch

## AOT Compatibility

COM interop calls (IContextMenu, IShellFolder) use `Marshal.GetObjectForIUnknown` which requires reflection. The `rd.xml` trimming descriptor preserves these types. The `[DynamicallyAccessedMembers]` attribute is applied to `ShellIntegrationService`.

Plain P/Invoke (SHGetFileInfo, SHFileOperation, Cloud Filter API) is fully AOT-safe.

The `WScript.Shell` COM usage in `CreateDesktopShortcutAsync` uses late-bound `dynamic` dispatch and is **not** AOT-compatible. For AOT builds, replace with direct `IShellLink` P/Invoke.

## Logging

Structured logs via Serilog are written to:
```
%LOCALAPPDATA%\FileExplorer\logs\log-yyyyMMdd.txt
```

## License

Copyright (c) FileExplorer contributors. All rights reserved.
