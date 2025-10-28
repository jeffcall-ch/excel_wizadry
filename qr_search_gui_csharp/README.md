# RQ GUI - Fast File Search Application

A modern Windows GUI application for the `rq` command-line file search utility. Provides a native Windows File Explorer-like interface for fast and powerful file searching with full search history management.

## 📦 Distribution

### Standalone Executable

The application is packaged as a **single, self-contained executable** that includes:
- ✅ .NET 8.0 Runtime (embedded)
- ✅ rq.exe search engine (embedded)
- ✅ All dependencies

**Location of shareable file:**
```
qr_search_gui_csharp\RqGui\bin\Release\net8.0-windows\win-x64\publish\RqGui.exe
```

**File size:** ~68.59 MB

### System Requirements

- **Operating System:** Windows 10/11 (x64)
- **No additional software required** - everything is embedded!

## 🚀 Usage

### Getting Started

1. **Launch** the application by double-clicking `RqGui.exe`
2. **Select a directory** to search:
   - Click **Browse** button, or
   - Type the path directly in the "Directory" field
3. **Enter search term** in "Search For" field (supports wildcards like `*.txt`)
4. **Press Enter** or click **Search** button

### Search Features

#### Basic Search
- Enter search term: `*.pdf` to find all PDF files
- Use wildcards: `report*.docx` to find files starting with "report"
- Default search: `*` (all files)

#### Advanced Filters

**Pattern Matching:**
- ☑️ **Glob Patterns** - Use wildcards (*, ?, [], {})
- ☑️ **Use Regex** - Regular expression matching
- ☑️ **Case Sensitive** - Exact case matching
- ☑️ **Include Hidden** - Show hidden files

**File Filters:**
- **Extensions:** `.txt,.pdf,.docx` (comma-separated)
- **Type:** `image`, `video`, `audio`, `document`, `archive`, `code`
- **Size:** `+100M` (larger than 100MB) or `-5K` (smaller than 5KB)
- **Depth:** Limit subdirectory depth (e.g., `3`)

**Date Filters:**
- **After:** Files modified after specified date
- **Before:** Files modified before specified date

**Performance Options:**
- **Threads:** Number of parallel threads (e.g., `8`)
- **Timeout:** Search timeout in milliseconds (e.g., `30000`)

### Search History

#### Using Search History
1. Click **"Recent Searches ▼"** button
2. **Click any search** to restore settings and execute immediately
3. **Hover over search** and click **🗑️ Delete** to remove it
4. **Click "Clear All History"** to remove all saved searches

#### What's Saved
Each search saves:
- Directory path
- Search term
- All filter checkboxes (glob, regex, case-sensitive, etc.)
- All filter values (extensions, file type, size, dates, etc.)
- Performance settings (depth, threads, timeout)

**History limit:** 50 most recent searches

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Enter` | Execute search (works from any field) |
| `Ctrl+O` | Open selected file |
| `Ctrl+Shift+O` | Open containing folder |
| `Ctrl+F` | Focus search box |
| `Escape` | Cancel active search |

### Results Actions

**Double-click** any result to open the file

**Right-click** on results for:
- **Open** - Opens file in default application
- **Open Containing Folder** - Opens File Explorer with file selected
- **Copy Full Path** - Copies full path to clipboard
- **Copy File Name** - Copies filename to clipboard

**Column Sorting:**
- Click column headers to sort
- Click again to reverse sort order
- Sortable columns: Name, Path, Size, Modified, Created

## 🛠️ Building from Source

### Prerequisites
- .NET 8.0 SDK
- Visual Studio Code or Visual Studio

### Build Commands

**Debug build:**
```powershell
cd qr_search_gui_csharp\RqGui
dotnet build -c Debug
```

**Release build:**
```powershell
cd qr_search_gui_csharp\RqGui
dotnet build -c Release
```

**Publish standalone executable:**
```powershell
cd qr_search_gui_csharp\RqGui
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

Output location: `bin\Release\net8.0-windows\win-x64\publish\RqGui.exe`

## 📁 Configuration & Data

### Settings Location
Application settings and search history are stored in:
```
%APPDATA%\RqGui\settings.json
```

### Log File
Debug logs are written to:
```
%APPDATA%\RqGui\rqgui.log
```

### Embedded rq.exe
The rq.exe search engine is extracted to:
```
%TEMP%\RqGui\rq.exe
```

You can also configure a custom rq.exe path via **Tools → Settings**.

## 🎯 Features Summary

- ✅ **Fast file search** powered by rq search engine
- ✅ **Press Enter anywhere** to execute search
- ✅ **Complete search history** with one-click restore & execute
- ✅ **Delete individual searches** or clear all history
- ✅ **Rich filtering options** (extensions, type, size, dates, depth)
- ✅ **Pattern matching** (glob patterns and regex)
- ✅ **Sortable results** with detailed file information
- ✅ **Context menu actions** for quick file operations
- ✅ **Fully standalone** - no installation required
- ✅ **Embedded dependencies** - .NET runtime and rq.exe included

## 📝 Tips

1. **Quick Re-search:** Use search history to instantly re-run complex searches
2. **Wildcards:** Use `*.{jpg,png,gif}` to search multiple extensions
3. **Size Filters:** Use `+1GB` for large files, `-100KB` for small files
4. **Performance:** Adjust thread count for faster searches on large directories
5. **Hidden Files:** Enable "Include Hidden" to search system and hidden files

## 🔗 Credits

- **rq search engine:** [github.com/seeyebe/rq](https://github.com/seeyebe/rq)
- **Framework:** .NET 8.0 Windows Forms

## 📄 License

This GUI application is provided as-is. The embedded `rq` utility maintains its original license.

---

**Version:** 1.0  
**Last Updated:** October 28, 2025  
**Platform:** Windows 10/11 x64
