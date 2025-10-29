# Shortcut Launcher - Taskbar Pinned WPF App

A Windows 11/10 taskbar-pinned application (.NET 8 WPF) for quick access to your favorite folders and files.

## 🎯 What This Does

Left-click a taskbar icon → Popup menu with all your shortcuts → Click to open!

Perfect for:
- Quick access to project folders
- Launching frequently used applications
- One-click navigation to network drives
- Organizing work/personal shortcuts

## 📂 Project Structure

```
ShortcutLauncher/
├── *.cs, *.xaml              # C# WPF application source code
├── ShortcutLauncher.csproj   # .NET 8 project file
├── ShortcutLauncher.sln      # Visual Studio solution
│
├── INDEX.md                  # 👈 START HERE - Project overview
├── SETUP_CHECKLIST.md        # ✅ Step-by-step setup guide
├── README.md                 # 📘 Complete user documentation
├── QUICKSTART.md             # ⚡ Quick build reference
├── PROJECT_STRUCTURE.md      # 🏗️ Code architecture details
├── ICON_SETUP.md             # 🎨 Icon creation help
│
├── build.bat                 # Quick build script
├── GenerateIcon.ps1          # Icon generator helper
├── appsettings.json          # Optional configuration
└── appicon-PLACEHOLDER.txt   # Icon setup instructions
```

## 🚀 Quick Start (First Time Users)

### Step 1: Create Icon (Required!)
```powershell
cd ShortcutLauncher
.\GenerateIcon.ps1
# This creates appicon.png
# Convert to .ico at: https://convertio.co/png-ico/
# Save as appicon.ico in this folder
```

### Step 2: Build
```powershell
.\build.bat
# OR:
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained
```

### Step 3: Pin to Taskbar
1. Go to: `bin\Release\net8.0-windows\win-x64\publish\`
2. Right-click `ShortcutLauncher.exe`
3. Click "Pin to taskbar"

### Step 4: Use It!
1. Left-click the taskbar icon
2. Click "Open Shortcuts Folder"
3. Add `.lnk` shortcut files to this folder
4. Click taskbar icon again - shortcuts appear!

## 📖 Documentation

**New user?** Follow this order:

1. **INDEX.md** - High-level overview (comprehensive summary)
2. **SETUP_CHECKLIST.md** - Detailed setup with checkboxes ✅
3. **QUICKSTART.md** - Fast reference for build commands
4. **README.md** - Complete feature documentation

**Developer?**

- **PROJECT_STRUCTURE.md** - Code architecture and details
- **MainWindow.xaml.cs** - Main application logic (well-commented)

**Having issues?**

- **ICON_SETUP.md** - Icon creation troubleshooting
- **README.md** - Troubleshooting section at the end

## ✨ Features

- ✅ Pin to taskbar with custom icon
- ✅ Left-click shows popup menu
- ✅ Lists all shortcuts from `%USERPROFILE%\Shortcuts_to_folders`
- ✅ Tooltips show shortcut targets
- ✅ Refresh, Open Folder, Exit menu items
- ✅ Ctrl+R hotkey to refresh
- ✅ Auto-creates folder if missing
- ✅ Handles 50+ shortcuts gracefully
- ✅ Optional configuration via JSON

## 🛠️ Requirements

- Windows 11 or Windows 10
- .NET 8 SDK (for building)
- .NET 8 Runtime (for running, unless built self-contained)

## 📦 Build Output

**Self-Contained (Recommended):**
- Single EXE: ~70 MB
- No .NET runtime required
- Portable - runs anywhere

**Framework-Dependent:**
- EXE + DLLs: ~200 KB
- Requires .NET 8 runtime installed
- Smaller file size

## 🎨 Customization

### Change Shortcuts Folder
Edit `appsettings.json`:
```json
{
  "ShortcutsFolder": "D:\\MyCustomFolder",
  "MaxItems": 100
}
```

### Change Icon
1. Replace `appicon.ico`
2. Rebuild
3. Unpin and re-pin to taskbar

## 🐛 Troubleshooting

| Issue | Solution |
|-------|----------|
| Default icon shows | Create/add `appicon.ico`, rebuild, re-pin |
| App won't start | Install .NET 8 or build self-contained |
| No shortcuts appear | Check folder path, ensure .lnk files exist |
| Menu doesn't show | Click taskbar icon, check Task Manager if running |

## 📜 License

Free to use and modify for any purpose.

## 🙌 Get Started

```powershell
# Clone/download the project
cd ShortcutLauncher

# Follow SETUP_CHECKLIST.md for detailed steps
# Or use quick start above for fast setup

# Questions? Check INDEX.md for comprehensive info
```

---

**Version:** 1.0  
**Framework:** .NET 8 WPF  
**Platform:** Windows x64  
**Created:** October 2025

**📖 Read Next:** [INDEX.md](ShortcutLauncher/INDEX.md) or [SETUP_CHECKLIST.md](ShortcutLauncher/SETUP_CHECKLIST.md)
