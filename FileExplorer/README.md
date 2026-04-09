# FileExplorer

WinUI 3 desktop file manager for Windows 11, built on .NET 8 and Windows App SDK.

## Solution structure

```
FileExplorer.sln
├── FileExplorer.App             # WinUI shell, startup, XAML views
├── FileExplorer.Core            # ViewModels, models, domain interfaces
├── FileExplorer.Infrastructure  # File system, shell, preview, cloud status services
└── FileExplorer.Tests           # xUnit tests
```

## Main features

- Tabbed file browsing with per-tab navigation history
- Split layout with folder tree, file list, preview pane, and details sections
- Native shell context menu integration (Explorer-style right-click commands)
- Rename, create, delete, move, copy, and conflict-aware file operations
- Undo support for reversible operations
- Live file refresh with debounced watcher behavior
- Search and filtering over file/folder listings
- OneDrive/Cloud status indicators via Cloud Filter APIs
- Keyboard-first workflow (Explorer-like shortcuts and focus behavior)
- Session persistence for tabs and view settings
- Structured logging and crash-report output for troubleshooting

## Prerequisites

### Runtime prerequisites (to run the app)

- Windows 11

### Build/publish prerequisites (to develop or deploy)

- Windows 11
- .NET SDK 8.x
- Windows SDK 10.0.22621.0+

Verify SDKs:

```powershell
dotnet --info
```

## Build and run (development)

```powershell
dotnet restore
dotnet build FileExplorer.sln -c Debug -r win-x64
dotnet run --project .\FileExplorer.App\FileExplorer.App.csproj -c Debug -r win-x64
```

If you see `WindowsAppSDK SelfContained requires a supported Windows architecture`, always pass an explicit runtime identifier (`-r win-x64`).

## Run tests

```powershell
dotnet test .\FileExplorer.Tests\FileExplorer.Tests.csproj
```

## Identity normalization rules (Project and Revision)

The Extracted Project and Extracted Revision columns are normalized before being shown and cached:

- Project names are matched against an allowlist CSV with fuzzy canonical matching.
- Revisions are shown only if they are in strict numeric format: `N.N`, `0.0`, `10.1`.

### Configure allowed project names

Default allowlist file:

- `Development\allowed_project_names.csv`

CSV format (single column, header optional):

```csv
Project
Thameside ERF
Walsall ERF
Medworth EfW CHP Facility
```

Matching behavior:

- Exact and fuzzy matching are both used.
- On a high-confidence fuzzy match, the displayed value is replaced by the canonical CSV value.
- If no high-confidence match is found, the project column is left blank.

Optional override:

- Set environment variable `FILEEXPLORER_ALLOWED_PROJECTS_CSV` to an absolute CSV path.

Revision behavior:

- Only `^\d+\.\d+$` is accepted.
- Any other revision form is cleared from display/cache.

## Reliable portable packaging (one-folder)

### Why this approach

For this app type (unpackaged WinUI 3 + Windows App SDK), strict single-file deployment has shown startup reliability problems in real-world scenarios due to resource resolution (`ms-appx`/PRI lookup) at runtime.

The most reliable portable strategy is therefore:

- Self-contained publish
- Non-single-file output
- Keep all runtime/resource files in one folder next to `FileExplorer.exe`

This repository now implements that strategy directly.

### What was implemented

- Publish profile: `FileExplorer.App/Properties/PublishProfiles/PortableFolder-x64.pubxml`
- Packaging script: `scripts/publish-portable.ps1`
- Publish safeguard in project file to ensure PRI files are copied into portable publish output

### Create the portable folder package

Default (Release + win-x64 + zip):

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-portable.ps1
```

Optional: skip zip creation

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-portable.ps1 -SkipZip
```

Optional: Debug build artifact

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-portable.ps1 -Configuration Debug
```

Script parameters:

- `-Configuration Release|Debug` (default `Release`)
- `-RuntimeIdentifier win-x64` (currently fixed to `win-x64`)
- `-SkipZip` (skip `.zip` output)

### Output

- Portable folder: `artifacts\portable-folder\win-x64\FileExplorer\`
- Optional zip: `artifacts\portable-folder\win-x64\FileExplorer-portable-win-x64.zip`

Run the app from:

- `artifacts\portable-folder\win-x64\FileExplorer\FileExplorer.exe`

### Deploy portable build to another machine

1. Build on your packaging machine using the script above.
2. Copy either:
	- the whole folder `artifacts\portable-folder\win-x64\FileExplorer\`, or
	- the zip `artifacts\portable-folder\win-x64\FileExplorer-portable-win-x64.zip`.
3. On target machine, if using zip, extract it fully to a normal writeable folder (do not run from inside zip).
4. Launch `FileExplorer.exe` from the extracted folder.

Recommended target folder examples:

- `%USERPROFILE%\Apps\FileExplorer\`
- `C:\Tools\FileExplorer\`

### Portable deployment update process

1. Close all running `FileExplorer.exe` instances.
2. Build a new portable package.
3. Replace the old folder contents with the new package contents.
4. Start `FileExplorer.exe` again.

### Portable packaging internals

The portable publish uses:

- `SelfContained=true`
- `WindowsPackageType=None`
- `WindowsAppSDKSelfContained=true`
- `PublishSingleFile=false`
- `PublishTrimmed=false`

During publish, the build includes an explicit step to place required PRI resources (including `resources.pri`) in the portable output folder to avoid missing-resource startup failures.

For maximum startup reliability, the packaging script then assembles the final portable folder from the app's runtime build output (`FileExplorer.App\bin\Release\...\win-x64`) after publish completes. This avoids known publish-layout regressions seen with unpackaged WinUI loose-XAML loading.

Required files checked by packaging script:

- `FileExplorer.exe`
- `resources.pri`
- `MainWindow.xaml`
- `Views\FileListView.xaml`

## Alternative publish profiles

- MSIX profile: `FileExplorer.App/Properties/PublishProfiles/MSIX-x64.pubxml`
- AOT profile: `FileExplorer.App/Properties/PublishProfiles/AOT-x64.pubxml`

Note: AOT/single-file options are available for experimentation, but the portable-folder path above is the recommended reliability default.

## Manual portable publish (without script)

If you need to run publish manually, use:

```powershell
dotnet publish .\FileExplorer.App\FileExplorer.App.csproj `
	-c Release `
	-r win-x64 `
	--self-contained true `
	/p:PublishProfile=Properties\PublishProfiles\PortableFolder-x64.pubxml `
	/p:PortableFolderPublish=true
```

Important: the script remains the recommended path because it also assembles the final artifact from known-good runtime output and validates required files.

## MSIX deployment (optional)

Use this when you want packaged installation behavior instead of portable one-folder deployment.

Publish command:

```powershell
dotnet publish .\FileExplorer.App\FileExplorer.App.csproj `
	-c Release `
	-r win-x64 `
	/p:PublishProfile=Properties\PublishProfiles\MSIX-x64.pubxml
```

Typical MSIX deployment flow:

1. Publish using the profile above.
2. Transfer produced MSIX bundle/package to target machine.
3. Install signing certificate on target machine if using a test/self-signed certificate.
4. Install package by double-clicking `.msix/.msixbundle` or via PowerShell `Add-AppxPackage`.

## AOT profile (experimental)

Publish command:

```powershell
dotnet publish .\FileExplorer.App\FileExplorer.App.csproj `
	-c Release `
	-r win-x64 `
	/p:PublishProfile=Properties\PublishProfiles\AOT-x64.pubxml
```

Use this profile for testing/perf experiments only. Validate startup and feature completeness before shipping.

## Deployment verification checklist

After deployment, verify:

1. `FileExplorer.exe` starts and stays running.
2. Folder navigation works from both tree and file-list flyouts.
3. New tab / new window actions work.
4. No new crash report appears after startup.

Quick startup verification from PowerShell:

```powershell
$exe = ".\artifacts\portable-folder\win-x64\FileExplorer\FileExplorer.exe"
Start-Process -FilePath $exe
Start-Sleep -Seconds 20
Get-Process FileExplorer -ErrorAction SilentlyContinue
```

## Deployment troubleshooting

### Build fails with architecture error

Symptom:

- `WindowsAppSDK SelfContained requires a supported Windows architecture`

Fix:

- Build/run/publish with `-r win-x64`.

### Portable app does not start

Check:

1. You are launching `FileExplorer.exe` from the extracted folder, not from inside zip.
2. Required runtime files are present next to the executable, especially `resources.pri` and XAML files.
3. Folder path is local and writeable.

### Packaging fails because files are locked

If another `FileExplorer.exe` is still running, packaging can fail during publish/resource generation. Close all app instances and run packaging again.

## Logs and crash reports

- Logs: `%LOCALAPPDATA%\FileExplorer\logs\`
- Managed crash reports: `%LOCALAPPDATA%\FileExplorer\crash-reports\`
- WER dumps: `%LOCALAPPDATA%\FileExplorer\crash-reports\wer-dumps\`

## Shortcut launcher flyout and startup prefetch

### Overview

The toolbar contains a **Shortcut Launcher** button that opens a cascading flyout menu. Each entry in the menu represents a `.lnk` shortcut from a designated folder on disk (the *shortcuts folder*). Clicking an entry navigates the active tab to the `.lnk` target. Subfolders inside the target are reachable via nested submenus without any disk I/O delay.

To make both the flyout and folder navigation instant, the app runs a multi-phase background warm-up immediately after the main window is shown.

### Shortcuts folder

The app resolves the shortcuts folder in this priority order:

1. **Environment variable** `FILEEXPLORER_SHORTCUTS_FOLDER` — if set and the path exists, it is used as-is. Environment variables inside the value are expanded.
2. **`%USERPROFILE%\Shortcuts_to_folders`** — used if that folder exists.
3. **Nearby folder auto-detect** — walks up from `AppContext.BaseDirectory` looking for `Shortcut_to_folders` or `Shortcuts_to_folders` up to 10 parent levels.
4. **Fallback** — `%USERPROFILE%\Shortcuts_to_folders` is returned regardless of whether it exists. The app creates it on first use.

`.lnk` files in the folder are resolved to their `TargetPath` at runtime. Only targets that exist on disk appear in the menu (up to `ShortcutMenuMaxItems = 50`).

**Adding .lnk files**: drop them into the shortcuts folder and restart the app. The flyout and prefetch are built once at startup; there is no file watcher on the shortcuts folder. The file-list tab *does* watch for changes if you have the shortcuts folder open in a tab, so a newly added `.lnk` appears there instantly.

### Startup sequence

The warm-up runs entirely on background threads after the window is displayed. It has three phases:

#### Phase 1 — shallow flyout cache (~10 ms)

Reads the immediate children (subfolders only) of every shortcut target in parallel. The result populates `_shortcutFolderCache` — an in-memory dictionary keyed by path. After phase 1 completes, the flyout is created on a low-priority dispatcher frame. Opening the flyout for the first time and expanding any first-level submenu costs zero disk I/O.

#### Phase 2 — deep flyout cache (~18 s for large trees)

Recursively walks every reachable subfolder of every shortcut target (up to `ShortcutMenuMaxDepth = 64` levels, `ShortcutMenuMaxSubfolders = 50` entries per node). This runs entirely on a background `Task.Run` thread. When it finishes, it updates `_shortcutFolderCache` with the full tree and then invalidates the already-built flyout XAML objects on a low-priority dispatcher frame so the next flyout open rebuilds from the complete cache. Expanding any submenu at any depth is now instant.

Phase 2 can take several seconds for deep or large OneDrive trees. It does not block the UI.

#### Phase 3 — directory snapshot prefetch (~350 ms)

Immediately after phase 1, a second background task pre-scans the *file contents* (not just subfolder names) of every shortcut target folder plus the first **eight** levels of subfolders (BFS). The results are stored in `IDirectorySnapshotCache` keyed by normalized path.

When the user navigates to a path that has a cached snapshot (`TryConsume`), `StreamLoadDirectoryAsync` skips all disk I/O and posts the cached `FileSystemEntry` list directly to the UI collection. Navigation to those folders appears instantaneous (0 ms scan, ~3 ms total to items shown). The snapshot is consumed exactly once — subsequent navigations to the same path perform a real disk scan so the view is never stale.

Typical numbers from DbgView:
```
Shortcut cache preload phase1 elapsed: 10 ms, roots=14
[SnapshotPrefetch] starting  paths=471
Shortcut cache preload total elapsed: 357 ms    ← snapshot prefetch fires here
[SnapshotPrefetch] done  scanned=471/471  ms=353
Shortcut cache preload phase2 elapsed: 18044 ms, cached=30342
```

```
[Navigate] snapshot cache HIT  165 items  path=...\Downloads
[Navigate] scan done  165 items  0 ms          ← 0 ms disk I/O
[Navigate] items shown  3 ms
```

### How the flyout builds submenus lazily

The flyout XAML is never fully pre-built for all levels at once. Building a `MenuFlyoutSubItem` tree for thousands of cache entries up-front would block the UI thread for seconds. Instead:

- The root menu entries are created once after phase 1 and reused across opens.
- Each `MenuFlyoutSubItem` gets a placeholder child. When the user hovers over a submenu (`PointerEntered`), its children are populated from `_shortcutFolderCache` on the UI thread. Because the cache is already in memory, this takes 0–1 ms per level.
- If the cache entry for a path is not yet populated (i.e. phase 2 has not reached it yet), the code falls back to a synchronous directory read capped at `ShortcutMenuMaxSubfolders` entries.

### Limits

| Constant | Value | Purpose |
|---|---|---|
| `ShortcutMenuMaxItems` | 50 | Max `.lnk` files shown in the root flyout |
| `ShortcutMenuMaxSubfolders` | 50 | Max subfolders shown per submenu node |
| `ShortcutMenuMaxDepth` | 64 | Max nesting depth before submenu expansion stops |
