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
