# RQ GUI

A modern Windows GUI application for the `rq` command-line file search utility. Provides a native Windows File Explorer-like interface for fast and powerful file searching.

## Features

- **Fast File Search**: Leverages the rq search engine for rapid file discovery
- **Flexible Filtering**: 
  - Glob patterns (wildcards: *, ?, [], {})
  - File extension filtering
  - File type filtering (image, video, audio, document, archive, code)
  - Size filtering (>, <)
  - Case-sensitive search
  - Hidden file inclusion
- **User-Friendly Interface**:
  - Windows 11 modern design
  - Sortable columns
  - Context menu for quick actions
  - Keyboard shortcuts
  - Progress indication
  - Search cancellation
- **Safe and Secure**: Command injection protection with proper argument handling
- **Performance**: Handles 10,000+ results efficiently with optimized UI updates

## Requirements

- **Windows 10/11** 
- **.NET 8.0 Runtime** ([Download](https://dotnet.microsoft.com/download/dotnet/8.0))
- **rq.exe** - Download from [github.com/seeyebe/rq/releases](https://github.com/seeyebe/rq/releases)

## Installation

### Quick Start
1. Download `rq.exe` from [rq releases](https://github.com/seeyebe/rq/releases)
2. Build this application:
   ```powershell
   cd RqGui
   dotnet build -c Release
   ```
3. Run `bin\Release\net8.0-windows\RqGui.exe`
4. Go to **Tools → Settings** and point to your `rq.exe` location

### Build from Source
```powershell
# Clone/navigate to the RqGui directory
cd RqGui

# Build release version
dotnet build -c Release

# Or publish as single file
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
```

## Usage

### Basic Search
1. Click **Browse** to select the directory to search
2. Enter your search term in **Search For**
3. Click **Search** or press **Enter**
4. Double-click results to open files

### Advanced Filtering

**Glob Patterns:**
- Check "Use Glob Patterns (*?[])"
- Examples:
  - `*.txt` - All text files
  - `report_202?.docx` - report_2020.docx, report_2021.docx, etc.
  - `*.{jpg,png,gif}` - All image files

**Extensions:**
- Enter comma-separated extensions: `jpg,png,pdf`

**File Types:**
- Select from dropdown: image, video, audio, document, archive, code, or Any

**Size Filtering:**
- Select `>` or `<` operator
- Enter size with unit: `100M`, `5K`, `1G`
- Examples:
  - `> 100M` - Files larger than 100 megabytes
  - `< 5K` - Files smaller than 5 kilobytes

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+F` | Focus search box |
| `Enter` | Start search (when in search controls) |
| `Escape` | Cancel active search |
| `Ctrl+O` | Open selected file |
| `Ctrl+Shift+O` | Open containing folder |

### Context Menu (Right-Click on Results)
- **Open** - Opens file in default application
- **Open Containing Folder** - Opens File Explorer with file selected
- **Copy Full Path** - Copies full path to clipboard
- **Copy File Name** - Copies filename to clipboard

### Sorting Results
- Click any column header to sort
- Click again to reverse sort order
- Columns: Name, Path, Size, Modified, Type

## Configuration

Settings are stored in: `%APPDATA%\RqGui\settings.json`

- `RqExePath`: Path to rq.exe
- `LastSearchPath`: Last searched directory (convenience feature)

You can manually edit this file or use **Tools → Settings** menu.

## Logs

Application logs are written to: `%APPDATA%\RqGui\rqgui.log`

If you encounter issues, check this file for error details.

## Troubleshooting

**"rq.exe not found"**
- Configure rq.exe location via Tools → Settings
- Or place rq.exe in your PATH environment variable
- Download from: https://github.com/seeyebe/rq/releases

**No results found**
- Check directory path is correct
- Try without filters first
- Verify search term spelling
- Check case-sensitive option if enabled
- Ensure rq.exe has permission to access the directory

**Search is slow**
- Use more specific search terms
- Enable glob patterns only when needed
- Use extension/type filters to narrow search
- Reduce max results if you only need a few files

**Files won't open**
- Verify file still exists (may have been deleted/moved after search)
- Check file associations in Windows
- Try "Open Containing Folder" instead

## Security

This application uses secure argument passing (`ArgumentList`) to prevent command injection attacks. User input is never concatenated into command strings.

## Architecture

- **Logger.cs**: Thread-safe file logging
- **SettingsManager.cs**: JSON-based settings persistence
- **FileResultItem.cs**: File metadata model
- **RqExecutor.cs**: Secure rq.exe process execution with cancellation support
- **SettingsForm.cs**: Settings dialog
- **Form1.cs**: Main application UI and logic

## Known Limitations

- Maximum 100,000 results per search (adjustable via Max Results setting)
- Network drives may be slower depending on connection speed
- Requires rq.exe to be downloaded separately

## License

[Your License Here]

## Credits

- Built with .NET 8.0 WinForms
- Uses [rq](https://github.com/seeyebe/rq) by seeyebe for file searching
- Modern Windows 11 inspired design

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Roadmap

Future enhancements planned:
- Date range filters (--after, --before)
- Search history dropdown
- Export results to CSV
- Dark mode theme
- File preview panel
- Saved search presets
