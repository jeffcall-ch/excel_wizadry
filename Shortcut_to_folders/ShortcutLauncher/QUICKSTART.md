# Quick Build & Deploy Guide

## Build the Application

### Option 1: Visual Studio 2022
1. Open `ShortcutLauncher.sln`
2. Press `Ctrl+Shift+B` to build
3. Executable will be in `bin\Release\net8.0-windows\`

### Option 2: Command Line (Quick Build)
```powershell
cd ShortcutLauncher
dotnet build -c Release
```

### Option 3: Publish Single EXE (Recommended)
```powershell
cd ShortcutLauncher
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true --self-contained true
```

Output: `bin\Release\net8.0-windows\win-x64\publish\ShortcutLauncher.exe`

## Icon Setup (Important!)

Before building, generate a custom icon:

```powershell
# Run this in the ShortcutLauncher directory
.\GenerateIcon.ps1
```

This creates `appicon.png`. Convert it to `appicon.ico` using:
- https://convertio.co/png-ico/
- Or use any ICO converter tool

Place `appicon.ico` in the ShortcutLauncher directory before building.

## Deploy & Pin

1. **Build/Publish** the app (see above)

2. **Navigate to output folder:**
   ```powershell
   cd bin\Release\net8.0-windows\win-x64\publish
   ```

3. **Pin to Taskbar:**
   - Right-click `ShortcutLauncher.exe`
   - Click "Pin to taskbar"

4. **Test:**
   - Click the taskbar icon
   - Menu should appear with shortcuts

## Add Shortcuts

1. Click the taskbar icon
2. Click "Open Shortcuts Folder"
3. Add `.lnk` files to this folder
4. Click taskbar icon again - shortcuts appear!

## Auto-Start (Optional)

To run on Windows startup:

```powershell
# Create startup shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ShortcutLauncher.lnk")
$Shortcut.TargetPath = "C:\Path\To\Your\ShortcutLauncher.exe"
$Shortcut.Save()
```

## Troubleshooting

### "Icon not showing"
- Ensure `appicon.ico` exists before building
- Rebuild after adding icon
- Unpin and re-pin the app

### "App won't start"
- Install .NET 8 Runtime: https://dotnet.microsoft.com/download/dotnet/8.0
- Or publish with `--self-contained true` (larger file, no runtime needed)

### "No shortcuts appearing"
- Ensure `.lnk` files exist in `%USERPROFILE%\Shortcuts_to_folders`
- Click "Refresh" in the menu
- Check folder path in `appsettings.json`

## One-Line Build & Open

```powershell
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained; explorer bin\Release\net8.0-windows\win-x64\publish
```

---

**Pro Tip:** Keep the published EXE in a permanent location (like `C:\Program Files\ShortcutLauncher\`) before pinning to taskbar. If you move it later, the pin will break.
