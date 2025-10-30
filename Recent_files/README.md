# Recent Launcher

A lightweight Windows desktop application for quickly accessing frequently used files and folders.

## Features

- **Single-file WinForms application** - All code in `Program.cs`
- **File & Folder bookmarks** - Add both files and folders to your quick access list
- **Persistent storage** - Bookmarks saved to `%APPDATA%\RecentLauncher\bookmarks.json`
- **Multiple ways to add items**:
  - Click "Add..." button and choose file or folder
  - Drag & drop files/folders from Windows Explorer
  - Keyboard shortcut: `Ctrl+N`
- **Quick access**:
  - Double-click to open
  - Click "Open" button
  - Press `Enter` key
  - Keyboard shortcut: `Ctrl+O`
- **Context menu** - Right-click for Open, Open Location, Rename, Remove
- **Keyboard shortcuts**:
  - `Enter` - Open selected item
  - `Delete` - Remove selected item
  - `F2` - Rename selected item
  - `Ctrl+N` - Add new item
  - `Ctrl+O` - Open selected item
  - `Ctrl+L` - Open location (parent folder)
  - `Ctrl+S` - Save bookmarks
- **Missing path detection** - Shows "(missing)" for items that no longer exist
- **Window position memory** - Remembers window size and position across sessions

## Requirements

- Windows OS
- .NET 8.0 SDK or Runtime

## Building & Running

```powershell
# Navigate to the directory
cd Recent_files

# Build the application
dotnet build

# Run the application
dotnet run

# Or publish as self-contained executable
dotnet publish -c Release -r win-x64 --self-contained
```

## Usage

1. **Adding bookmarks**:
   - Click "Add..." and choose File or Folder
   - Or drag & drop files/folders from Explorer onto the list

2. **Opening bookmarks**:
   - Double-click an item
   - Select and press Enter
   - Select and click "Open"

3. **Opening location**:
   - Select an item and click "Open Location"
   - Opens parent folder and selects the item

4. **Renaming**:
   - Select an item and press F2
   - Or right-click → Rename

5. **Removing**:
   - Select an item and press Delete
   - Or click "Remove" button
   - Or right-click → Remove

## Technical Details

- **Framework**: .NET 8.0 WinForms
- **Language**: C# 12
- **UI**: Windows Forms (no designer, code-only)
- **Persistence**: JSON via System.Text.Json
- **Data location**: `%APPDATA%\RecentLauncher\bookmarks.json`

## Architecture

The application consists of:
- `Program` class - Entry point
- `Bookmark` class - Data model for individual bookmarks
- `AppData` class - Container for bookmarks and window bounds
- `MainForm` class - Main application window with all UI and logic
- `Prompt` class - Helper for input dialogs

All contained in a single `Program.cs` file for easy distribution and maintenance.
