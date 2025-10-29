# Shortcut Launcher - Setup Checklist

Follow this checklist to get your Shortcut Launcher up and running!

## ☐ Prerequisites

- [ ] Windows 11 (or Windows 10)
- [ ] .NET 8 SDK installed
  - Check: Run `dotnet --version` in PowerShell
  - Install: https://dotnet.microsoft.com/download/dotnet/8.0
- [ ] Visual Studio 2022 (optional, for GUI development)

## ☐ Step 1: Get the Code

- [ ] Clone or download the ShortcutLauncher project
- [ ] Open PowerShell/Terminal in the `ShortcutLauncher` directory

## ☐ Step 2: Create Custom Icon (Important!)

Choose one method:

### Method A: Generate and Convert (Easiest)
- [ ] Run: `.\GenerateIcon.ps1`
- [ ] Open the created `appicon.png` in browser
- [ ] Visit: https://convertio.co/png-ico/
- [ ] Upload `appicon.png` and convert
- [ ] Download as `appicon.ico`
- [ ] Save `appicon.ico` in the ShortcutLauncher directory
- [ ] Delete `appicon-PLACEHOLDER.txt`

### Method B: Download Free Icon
- [ ] Visit: https://icons8.com or https://www.flaticon.com
- [ ] Search: "folder shortcut" or "launch"
- [ ] Download as .ico format (256x256)
- [ ] Save as `appicon.ico` in ShortcutLauncher directory
- [ ] Delete `appicon-PLACEHOLDER.txt`

### Method C: Use Existing Icon
- [ ] Copy any existing .ico file you like
- [ ] Rename to `appicon.ico`
- [ ] Place in ShortcutLauncher directory
- [ ] Delete `appicon-PLACEHOLDER.txt`

**Verify:**
- [ ] File `appicon.ico` exists in ShortcutLauncher directory
- [ ] File size > 0 bytes
- [ ] Can preview icon in Windows Explorer

## ☐ Step 3: Build the Application

Choose your preferred method:

### Option A: Use Build Script (Easiest)
```powershell
.\build.bat
```

### Option B: Manual PowerShell Commands
```powershell
dotnet restore
dotnet build -c Release
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained
```

### Option C: Visual Studio
- [ ] Open `ShortcutLauncher.sln`
- [ ] Set configuration to "Release"
- [ ] Build > Publish > Create profile for win-x64

**Verify:**
- [ ] Build completed without errors
- [ ] Executable exists: `bin\Release\net8.0-windows\win-x64\publish\ShortcutLauncher.exe`
- [ ] File size: ~60-80 MB (self-contained) or ~200 KB (framework-dependent)

## ☐ Step 4: Test the Application

- [ ] Navigate to publish folder
- [ ] Double-click `ShortcutLauncher.exe`
- [ ] Window should minimize to taskbar immediately
- [ ] Custom icon visible on taskbar? (If not, check Step 2)
- [ ] Click taskbar icon
- [ ] Menu appears with "Shortcuts folder not found" or shortcuts list
- [ ] Click "Open Shortcuts Folder" - folder opens?
- [ ] Click "Exit" - app closes?

**If test fails:**
- Check Windows Event Viewer for errors
- Ensure .NET 8 runtime installed
- Rebuild as self-contained (see Step 3)

## ☐ Step 5: Move to Permanent Location

**Important:** Pin to taskbar AFTER moving to permanent location!

- [ ] Create permanent folder (e.g., `C:\Program Files\ShortcutLauncher\`)
- [ ] Copy `ShortcutLauncher.exe` to permanent folder
- [ ] Copy `appsettings.json` (optional) to same folder
- [ ] Test: Run from new location

## ☐ Step 6: Pin to Taskbar

- [ ] Navigate to permanent executable location
- [ ] Right-click `ShortcutLauncher.exe`
- [ ] Select "Pin to taskbar"
- [ ] Custom icon visible on pinned taskbar icon?

**Verify:**
- [ ] Click pinned icon - menu appears?
- [ ] Icon looks correct (not default WPF icon)?

**If icon is wrong:**
- [ ] Unpin from taskbar
- [ ] Verify `appicon.ico` exists
- [ ] Rebuild app (Step 3)
- [ ] Re-pin to taskbar

## ☐ Step 7: Add Shortcuts

### Create Shortcuts Folder
- [ ] Click pinned taskbar icon
- [ ] Click "Open Shortcuts Folder"
- [ ] Folder opens: `C:\Users\<YourName>\Shortcuts_to_folders`

### Add Shortcuts (Choose method)

**Method A: Create Shortcut to Folder**
- [ ] Right-click any folder in File Explorer
- [ ] Send to > Desktop (create shortcut)
- [ ] Move shortcut from Desktop to Shortcuts_to_folders

**Method B: Drag & Drop**
- [ ] Open Shortcuts_to_folders folder
- [ ] Hold Alt key
- [ ] Drag any folder/file into Shortcuts_to_folders
- [ ] Windows creates shortcut automatically

**Method C: Manual Creation**
- [ ] Right-click in Shortcuts_to_folders folder
- [ ] New > Shortcut
- [ ] Browse to target folder/file
- [ ] Name the shortcut
- [ ] Click Finish

### Test Shortcuts
- [ ] Click pinned taskbar icon
- [ ] Shortcuts appear in menu?
- [ ] Click a shortcut - target opens?
- [ ] Click "Refresh" - menu updates?

## ☐ Step 8: Optional Configuration

### Customize Settings (appsettings.json)
- [ ] Edit `appsettings.json` in app directory
- [ ] Change `ShortcutsFolder` path (if desired)
- [ ] Change `MaxItems` limit (default 50)
- [ ] Save file
- [ ] Restart app to apply changes

### Auto-Start on Windows Login (Optional)
- [ ] Press `Win+R`
- [ ] Type: `shell:startup`
- [ ] Press Enter
- [ ] Create shortcut to `ShortcutLauncher.exe` in Startup folder

**OR use PowerShell:**
```powershell
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ShortcutLauncher.lnk")
$Shortcut.TargetPath = "C:\Program Files\ShortcutLauncher\ShortcutLauncher.exe"
$Shortcut.Save()
```

- [ ] Log out and log back in
- [ ] App starts automatically?
- [ ] Icon appears on taskbar?

## ☐ Step 9: Final Verification

- [ ] ✅ Custom icon shows on taskbar
- [ ] ✅ Left-click shows menu instantly
- [ ] ✅ Shortcuts open correct targets
- [ ] ✅ "Open Shortcuts Folder" works
- [ ] ✅ "Refresh" updates menu
- [ ] ✅ "Exit" closes app cleanly
- [ ] ✅ Ctrl+R refreshes menu (when open)
- [ ] ✅ Multiple shortcuts display correctly
- [ ] ✅ Menu closes when clicking elsewhere

## ✅ Complete!

Your Shortcut Launcher is now ready to use!

## 🎯 Quick Reference

**Add new shortcut:**
1. Click taskbar icon
2. Open Shortcuts Folder
3. Add .lnk file
4. Click Refresh (or close/reopen menu)

**Change folder location:**
1. Edit appsettings.json
2. Set "ShortcutsFolder" to new path
3. Restart app

**Update icon:**
1. Replace appicon.ico
2. Rebuild: `dotnet build -c Release`
3. Unpin and re-pin to taskbar

## 🆘 Troubleshooting

| Problem | Solution |
|---------|----------|
| Default icon on taskbar | Rebuild with appicon.ico, unpin/re-pin |
| App won't start | Install .NET 8 runtime or rebuild as self-contained |
| Menu doesn't appear | Check if app is running (Task Manager) |
| Shortcuts not listed | Verify folder path, check .lnk files exist |
| "Access denied" error | Run as Administrator once, then normal |

---

**Need Help?** Check:
- README.md - Full documentation
- QUICKSTART.md - Build commands
- PROJECT_STRUCTURE.md - Architecture details
- ICON_SETUP.md - Icon creation help

**Version:** 1.0  
**Last Updated:** October 2025
