# Troubleshooting Guide - Shortcut Launcher

Common issues and solutions for the Shortcut Launcher app.

---

## 🔧 Build Issues

### ❌ "SDK not found" or "dotnet command not recognized"

**Cause:** .NET 8 SDK not installed or not in PATH

**Solution:**
```powershell
# Check if .NET is installed
dotnet --version

# If not found, download and install .NET 8 SDK:
# https://dotnet.microsoft.com/download/dotnet/8.0
```

### ❌ "The current .NET SDK does not support targeting .NET 8.0"

**Cause:** Older .NET SDK version installed

**Solution:**
```powershell
# Check version
dotnet --version
# Should be 8.0.x or higher

# Update to .NET 8:
# https://dotnet.microsoft.com/download/dotnet/8.0
```

### ❌ Build succeeds but EXE not found

**Cause:** Looking in wrong folder

**Solution:**
```powershell
# Check exact output path
dotnet build -c Release

# Output is in:
# bin\Release\net8.0-windows\ShortcutLauncher.exe

# For published:
# bin\Release\net8.0-windows\win-x64\publish\ShortcutLauncher.exe
```

### ❌ "IWshRuntimeLibrary could not be found"

**Cause:** COM reference registration issue

**Solution:**
```powershell
# Clean and rebuild
dotnet clean
dotnet restore
dotnet build -c Release

# If still fails, check Windows Script Host is enabled:
# Press Win+R, type: wscript
# Should open a dialog (close it)
```

### ❌ "appicon.ico not found" during build

**Cause:** Icon file missing or misnamed

**Solution:**
```powershell
# Check if file exists
Test-Path .\appicon.ico

# If False, create the icon (see ICON_SETUP.md)
# Temporary workaround: remove Icon property from .csproj

# Edit ShortcutLauncher.csproj, comment out:
# <ApplicationIcon>appicon.ico</ApplicationIcon>
```

---

## 🚀 Runtime Issues

### ❌ App won't start / Nothing happens when double-clicking

**Cause 1:** .NET 8 Runtime not installed (framework-dependent build)

**Solution:**
```powershell
# Check if runtime installed
dotnet --list-runtimes

# Should show: Microsoft.WindowsDesktop.App 8.0.x

# If missing, install .NET 8 Runtime:
# https://dotnet.microsoft.com/download/dotnet/8.0/runtime

# OR rebuild as self-contained:
dotnet publish -c Release -r win-x64 --self-contained
```

**Cause 2:** App is already running (single instance)

**Solution:**
```powershell
# Check Task Manager (Ctrl+Shift+Esc)
# Look for: ShortcutLauncher.exe
# End task if found, then restart
```

**Cause 3:** Crash on startup

**Solution:**
```powershell
# Check Windows Event Viewer
# Win+R → eventvwr → Windows Logs → Application
# Look for .NET Runtime errors

# Or run from command line to see errors:
cd bin\Release\net8.0-windows\win-x64\publish
.\ShortcutLauncher.exe
# Error messages will appear in console
```

### ❌ App starts but no taskbar icon appears

**Cause:** Window hidden but not showing in taskbar

**Solution:**
Check `MainWindow.xaml`:
```xml
<!-- Ensure this is set: -->
ShowInTaskbar="True"
```

Rebuild after fixing.

### ❌ App minimizes immediately, can't interact

**Cause:** Normal behavior - it's a taskbar app!

**Solution:**
```
This is expected behavior:
1. App starts
2. Minimizes to taskbar
3. Click taskbar icon to show menu
```

---

## 🎨 Icon Issues

### ❌ Default/generic icon shows in taskbar when pinned

**Cause:** `appicon.ico` missing or not embedded in build

**Solution:**
```powershell
# 1. Verify icon exists
Test-Path .\appicon.ico

# 2. Check icon is in .csproj
# Open ShortcutLauncher.csproj
# Should have: <ApplicationIcon>appicon.ico</ApplicationIcon>

# 3. Clean and rebuild
dotnet clean
dotnet build -c Release

# 4. Check EXE has icon
# Right-click bin\Release\...\ShortcutLauncher.exe
# Properties → Should show custom icon

# 5. Unpin from taskbar
# Right-click taskbar icon → Unpin

# 6. Re-pin to taskbar
# Navigate to EXE, right-click → Pin to taskbar
```

### ❌ Icon looks blurry or low quality

**Cause:** Icon not multi-resolution or too small

**Solution:**
```
Create icon with multiple resolutions:
- 16x16
- 32x32
- 48x48
- 256x256

Use online converter with "multi-size" option:
https://convertio.co/png-ico/

Check "Create multi-resolution icon"
```

### ❌ Can't create icon / conversion fails

**Cause:** Various issues with icon creation

**Solution:**
See [ICON_ALTERNATIVES.md](ICON_ALTERNATIVES.md) for 10 different methods.

**Quick workaround:**
```powershell
# Download a free icon from:
# https://icons8.com/icons/set/folder

# Search "folder", download as ICO, save as appicon.ico
```

---

## 📁 Shortcuts Issues

### ❌ "No shortcuts found" even though .lnk files exist

**Cause 1:** Wrong folder path

**Solution:**
```powershell
# Check where app is looking
# Default: %USERPROFILE%\Shortcuts_to_folders

# Verify path:
$env:USERPROFILE
# Example: C:\Users\YourName

# Full path should be:
# C:\Users\YourName\Shortcuts_to_folders

# Check if folder exists and has .lnk files:
Get-ChildItem "$env:USERPROFILE\Shortcuts_to_folders" -Filter *.lnk
```

**Cause 2:** Custom path in appsettings.json

**Solution:**
```powershell
# Check appsettings.json
cat .\appsettings.json

# Verify ShortcutsFolder path
# Ensure it exists and has .lnk files
```

### ❌ Shortcuts appear but clicking does nothing

**Cause:** Invalid .lnk file or target missing

**Solution:**
```powershell
# Test shortcut manually:
# 1. Open %USERPROFILE%\Shortcuts_to_folders
# 2. Double-click the .lnk file
# 3. Does it work?

# If not, recreate the shortcut:
# Right-click target folder → Send to → Desktop
# Move shortcut from Desktop to Shortcuts_to_folders
```

### ❌ "Access denied" when clicking shortcut

**Cause:** Target requires elevation or permissions

**Solution:**
```
Option 1: Run app as Administrator (once)
- Right-click ShortcutLauncher.exe
- Run as administrator
- Try shortcut again

Option 2: Fix shortcut permissions
- Ensure target folder is accessible
- Check network drive is connected
- Verify you have read permissions
```

### ❌ Only some shortcuts appear (not all)

**Cause:** MaxItems limit reached

**Solution:**
```json
// Edit appsettings.json
{
  "ShortcutsFolder": "%USERPROFILE%\\Shortcuts_to_folders",
  "MaxItems": 100  // Increase this
}

// Restart app
```

---

## 🖱️ Menu Issues

### ❌ Menu doesn't appear when clicking taskbar icon

**Cause 1:** App not actually running

**Solution:**
```
Check Task Manager (Ctrl+Shift+Esc)
- Look for ShortcutLauncher.exe
- If not found, start the app
```

**Cause 2:** Menu appearing off-screen

**Solution:**
```
Try clicking different areas of the taskbar icon
Or restart the app:
1. Task Manager → End ShortcutLauncher.exe
2. Click taskbar icon again
```

**Cause 3:** Crash in menu building code

**Solution:**
```powershell
# Run from command line to see errors
cd bin\Release\net8.0-windows\win-x64\publish
.\ShortcutLauncher.exe

# Click taskbar icon
# Errors will appear in console
```

### ❌ Menu appears in wrong location

**Cause:** Cursor position detection issue

**Solution:**
This is expected behavior - menu appears near cursor.

To change, edit `MainWindow.xaml.cs`:
```csharp
// Find ShowMenuAtCursor() method
// Adjust positioning logic
```

### ❌ Menu closes immediately

**Cause:** Focus loss or mouse click outside

**Solution:**
```
Normal behavior when:
- Clicking outside menu
- Pressing Esc
- Clicking another window

If closing too fast:
- Check mouse sensitivity
- Ensure not clicking outside accidentally
```

### ❌ Ctrl+R hotkey doesn't work

**Cause:** Menu not open or focus issue

**Solution:**
```
Ctrl+R only works when menu is open:
1. Click taskbar icon (menu opens)
2. Press Ctrl+R (menu refreshes)

If still doesn't work:
- Ensure menu has focus
- Click inside menu first
- Then press Ctrl+R
```

---

## ⚙️ Configuration Issues

### ❌ appsettings.json changes ignored

**Cause:** App caching old config or wrong location

**Solution:**
```powershell
# Ensure appsettings.json is next to EXE
# For published app:
# bin\Release\net8.0-windows\win-x64\publish\appsettings.json

# Check file is there:
Get-ChildItem bin\Release\net8.0-windows\win-x64\publish\appsettings.json

# Restart app after changing config
# (Close via "Exit" menu, then reopen)
```

### ❌ Environment variables not expanding

**Cause:** Incorrect syntax in JSON

**Solution:**
```json
// CORRECT:
{
  "ShortcutsFolder": "%USERPROFILE%\\Shortcuts_to_folders"
}

// WRONG (single backslash):
{
  "ShortcutsFolder": "%USERPROFILE%\Shortcuts_to_folders"
}

// JSON requires double backslashes
```

---

## 🔄 Update/Reinstall Issues

### ❌ Pinned icon stops working after moving EXE

**Cause:** Taskbar pin points to old location

**Solution:**
```
1. Unpin from taskbar (right-click → Unpin)
2. Move EXE to permanent location
3. Re-pin to taskbar
```

### ❌ Changes not showing after rebuild

**Cause:** Running old version from cache

**Solution:**
```powershell
# 1. Close app completely
# Task Manager → End ShortcutLauncher.exe

# 2. Clean build
dotnet clean
dotnet build -c Release

# 3. Unpin and re-pin to taskbar
```

---

## 🪟 Windows-Specific Issues

### ❌ "Windows protected your PC" SmartScreen warning

**Cause:** Unsigned executable

**Solution:**
```
Click "More info"
Click "Run anyway"

For distribution:
- Code sign the EXE (requires certificate)
- Or build with ClickOnce deployment
```

### ❌ Antivirus blocks or quarantines EXE

**Cause:** False positive (unsigned EXE)

**Solution:**
```
Add exception in your antivirus:
- Windows Defender: Settings → Virus & threat protection
  → Manage settings → Add exclusion
- Add: ShortcutLauncher.exe

For long-term:
- Code sign the executable
```

### ❌ Can't pin to taskbar (option greyed out)

**Cause:** Windows policy or EXE in temp location

**Solution:**
```
1. Move EXE to permanent location:
   - C:\Program Files\ShortcutLauncher\
   - C:\Users\YourName\Apps\

2. Right-click EXE again → Pin to taskbar

3. If still greyed, run as admin once:
   - Right-click → Run as administrator
   - Close app
   - Try pinning again
```

---

## 🧪 Testing/Debugging

### Enable Verbose Logging

Add debug output to `MainWindow.xaml.cs`:

```csharp
// In BuildMenu() method
Debug.WriteLine($"Shortcuts folder: {_shortcutsFolder}");
Debug.WriteLine($"Folder exists: {Directory.Exists(_shortcutsFolder)}");

// View output:
// Run from Visual Studio (F5)
// Output appears in Debug console
```

### Test Without Pinning

```powershell
# Run directly without pinning
cd bin\Release\net8.0-windows\win-x64\publish
.\ShortcutLauncher.exe

# Window should appear in taskbar
# Click to test menu
```

### Check Event Logs

```powershell
# Open Event Viewer
eventvwr

# Navigate to:
# Windows Logs → Application

# Filter for:
# Source: .NET Runtime
# Look for errors from ShortcutLauncher
```

---

## 📞 Still Having Issues?

### Gather Information:

```powershell
# System info
dotnet --info

# File locations
Get-ChildItem bin\Release\net8.0-windows\win-x64\publish

# Shortcuts folder
Get-ChildItem "$env:USERPROFILE\Shortcuts_to_folders"

# Running processes
Get-Process ShortcutLauncher -ErrorAction SilentlyContinue
```

### Reset Everything:

```powershell
# 1. Unpin from taskbar
# 2. Delete build outputs
Remove-Item -Recurse -Force bin, obj

# 3. Clean build
dotnet restore
dotnet build -c Release
dotnet publish -c Release -r win-x64 --self-contained

# 4. Pin fresh build to taskbar
```

### Check Documentation:

- **README.md** - Full feature docs
- **SETUP_CHECKLIST.md** - Step-by-step setup
- **ICON_SETUP.md** - Icon creation help
- **PROJECT_STRUCTURE.md** - Code architecture

---

## ✅ Common Success Checklist

If app works perfectly, you should see:

- [ ] Custom icon in taskbar when pinned
- [ ] Left-click shows menu instantly
- [ ] All .lnk files from Shortcuts_to_folders listed
- [ ] Clicking shortcut opens target
- [ ] "Refresh" rebuilds menu
- [ ] "Open Shortcuts Folder" opens folder
- [ ] "Exit" closes cleanly
- [ ] Ctrl+R refreshes menu
- [ ] Menu closes when clicking outside

If any item fails, see relevant section above.

---

**Last Updated:** October 2025  
**Version:** 1.0
