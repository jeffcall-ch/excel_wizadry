# Shortcut Launcher - Taskbar Pinned WPF App

A Windows 11 WPF application (.NET 8) that provides quick access to shortcuts from a custom folder via a taskbar-pinned icon.

## 📦 Ready-to-Use Executable

**Location:** `bin\Release\net8.0-windows\win-x64\publish\ShortcutLauncher.exe`

- **Size:** ~155 MB (self-contained with .NET 8 runtime)
- **Portable:** Copy this single file to any Windows 10/11 x64 machine
- **No installation required:** Just run the .exe file
- **No dependencies:** .NET runtime is included

## Features

✅ **Taskbar Integration**
- Pin to taskbar with custom icon
- Left-click to show shortcuts menu
- No tray icon - pure taskbar app

✅ **Shortcuts Management**
- Automatically lists all `.lnk` files from `%USERPROFILE%\Shortcuts_to_folders`
- Alphabetically sorted
- Tooltips show target paths
- Dynamic rebuild on each display
- Handles up to 50 shortcuts (configurable)
- Separators every 10 items for readability

✅ **Menu Features**
- Click any shortcut to open its target
- **Refresh** - Rebuild the shortcuts list (or press Ctrl+R)
- **Open Shortcuts Folder** - Manage your shortcuts
- **Exit** - Close the application

✅ **Configuration**
- Optional `appsettings.json` for custom folder path and max items
- Falls back to sensible defaults if config is missing
- Creates shortcuts folder automatically if it doesn't exist

## Requirements

- Windows 11 (or Windows 10)
- .NET 8 Runtime (or build as self-contained)

## Building the Application

### Option 1: Visual Studio 2022

1. Open `ShortcutLauncher.sln` in Visual Studio 2022
2. Build the solution (Ctrl+Shift+B)
3. The executable will be in `bin\Debug\net8.0-windows\` or `bin\Release\net8.0-windows\`

### Option 2: Command Line

```powershell
# Debug build
dotnet build

# Release build
dotnet build -c Release
```

### Option 3: Publish Single Executable

To create a single, self-contained executable:

```powershell
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true
```

The output will be in `bin\Release\net8.0-windows\win-x64\publish\`

## Installation & Setup

### 1. Add Custom Icon

Replace `appicon.ico` with your own icon file (recommended: 256x256 pixels, .ico format).

The default placeholder icon will work but you may want to customize it.

### 2. Build/Publish

Use one of the build methods above.

### 3. Pin to Taskbar

1. Navigate to the published executable location
2. Right-click `ShortcutLauncher.exe`
3. Select **"Pin to taskbar"**

The custom icon from `appicon.ico` will appear on the taskbar.

### 4. Add Shortcuts

1. Left-click the taskbar icon
2. Click **"Open Shortcuts Folder"**
3. Add `.lnk` files to this folder
4. Click the taskbar icon again - your shortcuts will appear!

## Usage

### Basic Usage

1. **Left-click** the taskbar icon
2. A popup menu appears listing all shortcuts
3. Click any shortcut to open its target
4. Menu closes automatically

### Keyboard Shortcuts

- **Ctrl+R** - Refresh the shortcuts list while menu is open

### Menu Options

- **🔄 Refresh** - Reload shortcuts from folder
- **📁 Open Shortcuts Folder** - Open the folder to manage shortcuts
- **❌ Exit** - Close the application

### Creating Shortcuts

You can create shortcuts in several ways:

**Method 1: Drag & Drop**
- Open the shortcuts folder (via the app menu)
- Drag any folder or file to the folder while holding Alt
- Windows will create a shortcut

**Method 2: Right-click**
- Right-click any folder/file
- Send to → Desktop (creates shortcut)
- Move the shortcut from Desktop to the Shortcuts_to_folders folder

**Method 3: Manual**
- Right-click in Shortcuts_to_folders
- New → Shortcut
- Browse to the target location
- Name the shortcut

## Configuration

### appsettings.json

Customize the shortcuts folder path or maximum items:

```json
{
  "ShortcutsFolder": "%USERPROFILE%\\Shortcuts_to_folders",
  "MaxItems": 50
}
```

**Parameters:**
- `ShortcutsFolder` - Path to the shortcuts folder (supports environment variables)
- `MaxItems` - Maximum number of shortcuts to display (default: 50)

### Default Values

If `appsettings.json` is missing, the app uses:
- Folder: `C:\Users\<YourUsername>\Shortcuts_to_folders`
- Max Items: 50

## Optional: Autostart

To run the app automatically on Windows startup:

1. Press `Win+R`
2. Type `shell:startup` and press Enter
3. Create a shortcut to `ShortcutLauncher.exe` in this folder

## Troubleshooting

### Icon Not Showing When Pinned

- Ensure `appicon.ico` exists in the same folder as the executable
- Rebuild the project after adding/changing the icon
- Unpin and re-pin the application

### Shortcuts Not Appearing

- Check that `.lnk` files exist in the shortcuts folder
- Click "Refresh" in the menu
- Verify the folder path in `appsettings.json`

### Application Won't Start

- Ensure .NET 8 runtime is installed
- Or publish as self-contained (see build instructions)
- Check Windows Event Viewer for error details

### Menu Appears in Wrong Location

- This is normal behavior - the menu appears near the cursor
- Click the taskbar icon again if needed

## Technical Details

**Framework:** .NET 8 WPF  
**Language:** C#  
**Platform:** Windows x64  
**Dependencies:**
- Microsoft.Extensions.Configuration.Json (for config)
- System.Drawing.Common (for icon handling)
- COM Interop with WScript.Shell (for shortcut target resolution)

**Window Configuration:**
- Visible window with buttons on startup
- Minimizes to taskbar when requested
- Left-click taskbar icon shows shortcuts menu
- Window close button prompts: minimize or exit

## Icon Configuration Notes

**Important for future reference:**

The application icon (`appicon.ico`) must be embedded as a WPF **Resource**, not just copied to output.

### Correct .csproj Configuration:
```xml
<ItemGroup>
  <Resource Include="appicon.ico" />
</ItemGroup>
```

### ❌ Incorrect (causes "Cannot locate resource" crash):
```xml
<ItemGroup>
  <None Update="appicon.ico">
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>
```

### XAML Reference:
- For MainWindow: Remove `Icon="appicon.ico"` if it causes issues during development
- The icon in `.csproj` `<ApplicationIcon>appicon.ico</ApplicationIcon>` sets the EXE icon (taskbar/file explorer)
- The XAML `Icon` property requires the resource to be embedded

### Why This Matters:
- WPF looks for icons in the compiled assembly resources
- `<Resource Include>` embeds the file into the DLL/EXE
- `<None Update>` only copies to output directory (file system)
- Single-file publish bundles resources but not loose files

## License

This is a utility application. Use and modify as needed for your personal or commercial projects.

## Credits

Created as a taskbar-pinned shortcut launcher for Windows 11.

---

**Version:** 1.0  
**Last Updated:** October 2025
