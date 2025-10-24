# RQ GUI Development Action Plan

## Project Overview
Build a Windows 11 C# WinForms application that provides a File Explorer-like GUI for the `rq` command-line file search utility. The application should hide the command-line interface and provide a native Windows experience.

**Target Framework:** .NET 6.0 or later (required for modern APIs)

**RQ Command Reference:** https://github.com/seeyebe/rq

---

## Step 1: Project Setup and Configuration

**Action:** Create a new C# WinForms project in VS Code with proper configuration.

```bash
# In VS Code terminal
dotnet new winforms -n RqGui -f net6.0-windows
cd RqGui
code .
```

**Add Application Icon:**
Create or obtain an `icon.ico` file and add to project root.

**Edit `RqGui.csproj`:**
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net6.0-windows</TargetFramework>
    <Nullable>enable</Nullable>
    <UseWindowsForms>true</UseWindowsForms>
    <ApplicationIcon>icon.ico</ApplicationIcon>
  </PropertyGroup>
  
  <ItemGroup>
    <Content Include="icon.ico" />
  </ItemGroup>
</Project>
```

**Verify:**
- Project structure created
- `.csproj` file exists with correct framework
- `Form1.cs` and `Program.cs` present
- Icon file included

**Test:** Run `dotnet build` to ensure project compiles.

---

## Step 2: Design Main Form Layout

**Action:** Design the main UI with File Explorer aesthetics.

**Components to add:**
1. **Top Panel (Search Controls)**
   - Label: "Search Directory"
   - TextBox: `txtSearchPath` (TabIndex: 0)
   - Button: `btnBrowseDirectory` ("Browse...", TabIndex: 1)
   - Label: "Search For"
   - TextBox: `txtSearchTerm` (TabIndex: 2)
   - Button: `btnSearch` ("Search", TabIndex: 3)
   - Button: `btnCancel` ("Cancel", initially disabled, TabIndex: 4)

2. **Middle Panel (Filter Options)**
   - GroupBox: "Filters"
   - CheckBox: `chkGlob` ("Use Glob Patterns (*?[])", TabIndex: 5)
   - CheckBox: `chkCaseSensitive` ("Case Sensitive", TabIndex: 6)
   - CheckBox: `chkIncludeHidden` ("Include Hidden Files", TabIndex: 7)
   - Label: "File Extensions (comma-separated)"
   - TextBox: `txtExtensions` (placeholder: "jpg,png,pdf", TabIndex: 8)
   - Label: "File Type"
   - ComboBox: `cmbFileType` (items: "Any", "image", "video", "audio", "document", TabIndex: 9)
   - Label: "File Size"
   - ComboBox: `cmbSizeOperator` (items: "Any", ">", "<", TabIndex: 10)
   - TextBox: `txtSizeValue` (placeholder: "100M", TabIndex: 11)
   - Label: "Max Results"
   - NumericUpDown: `numMaxResults` (default: 1000, min: 1, max: 100000, TabIndex: 12)

3. **Bottom Panel (Results)**
   - ListView: `lvResults` 
     - View: Details
     - FullRowSelect: true
     - GridLines: true
     - Columns: Name, Path, Size, Modified, Type
     - Sorting: Enabled
     - TabIndex: 13
   - ProgressBar: `progressBar1` (Style: Marquee, initially hidden)

4. **Status Bar**
   - StatusStrip with label showing result count and status

**Pseudo Layout:**
```
┌─────────────────────────────────────────────────────────┐
│ Search Directory: [C:\Users\...  ] [Browse]            │
│ Search For: [*.txt        ] [Search] [Cancel]          │
├─────────────────────────────────────────────────────────┤
│ Filters:                                                │
│ ☐ Glob  ☐ Case Sensitive  ☐ Hidden                    │
│ Extensions: [jpg,png] Type: [Any▼] Size: [>][100M]    │
│ Max Results: [1000]                                     │
├─────────────────────────────────────────────────────────┤
│ Name          │ Path      │ Size │ Modified│ Type      │
│ document.txt  │ C:\Docs   │ 12KB │ 1/1/25  │ .txt      │
│ ...           │           │      │         │           │
│ [▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓] (progress bar - hidden by default) │
└─────────────────────────────────────────────────────────┘
│ Found 42 files                                          │
└─────────────────────────────────────────────────────────┘
```

**Test:** 
- Run app, verify all controls are visible and properly aligned
- Press Tab key repeatedly, verify tab order flows logically
- Verify controls are accessible and properly labeled

---

## Step 3: Implement Settings Management

**Action:** Create a settings manager to store and retrieve `rq.exe` path.

**Create `SettingsManager.cs`:**
```csharp
using System;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace RqGui
{
    public class AppSettings
    {
        public string RqExePath { get; set; } = "";
        public string LastSearchPath { get; set; } = "";
    }

    public static class SettingsManager
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RqGui",
            "settings.json"
        );

        public static AppSettings LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                // Log error but continue with defaults
                Logger.LogError($"Failed to load settings: {ex.Message}");
            }
            return new AppSettings();
        }

        public static bool SaveSettings(AppSettings settings)
        {
            try
            {
                string? directory = Path.GetDirectoryName(SettingsPath);
                if (directory != null)
                {
                    Directory.CreateDirectory(directory);
                }
                
                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to save settings: {ex.Message}");
                MessageBox.Show(
                    $"Failed to save settings: {ex.Message}\n\nSettings will not be persisted.",
                    "Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return false;
            }
        }
    }
}
```

**Create `SettingsForm.cs`:**
```csharp
using System;
using System.IO;
using System.Windows.Forms;

namespace RqGui
{
    public partial class SettingsForm : Form
    {
        private TextBox txtRqPath;
        private Button btnBrowse;
        private Button btnSave;
        private Button btnCancel;
        private Label lblRqPath;
        
        public SettingsForm()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }
        
        private void InitializeComponent()
        {
            this.Text = "Settings";
            this.Size = new System.Drawing.Size(500, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            
            // Label
            lblRqPath = new Label
            {
                Text = "RQ.exe Location:",
                Location = new System.Drawing.Point(10, 20),
                Size = new System.Drawing.Size(100, 20)
            };
            
            // TextBox
            txtRqPath = new TextBox
            {
                Location = new System.Drawing.Point(120, 18),
                Size = new System.Drawing.Size(270, 20),
                TabIndex = 0
            };
            
            // Browse button
            btnBrowse = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(400, 16),
                Size = new System.Drawing.Size(75, 25),
                TabIndex = 1
            };
            btnBrowse.Click += BtnBrowse_Click;
            
            // Save button
            btnSave = new Button
            {
                Text = "Save",
                Location = new System.Drawing.Point(300, 70),
                Size = new System.Drawing.Size(75, 25),
                TabIndex = 2
            };
            btnSave.Click += BtnSave_Click;
            
            // Cancel button
            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(400, 70),
                Size = new System.Drawing.Size(75, 25),
                TabIndex = 3,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.Click += (s, e) => this.Close();
            
            this.CancelButton = btnCancel;
            
            this.Controls.Add(lblRqPath);
            this.Controls.Add(txtRqPath);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }
        
        private void LoadCurrentSettings()
        {
            var settings = SettingsManager.LoadSettings();
            txtRqPath.Text = settings.RqExePath;
        }
        
        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*";
                dialog.Title = "Select rq.exe";
                
                if (!string.IsNullOrWhiteSpace(txtRqPath.Text) && File.Exists(txtRqPath.Text))
                {
                    dialog.InitialDirectory = Path.GetDirectoryName(txtRqPath.Text);
                    dialog.FileName = Path.GetFileName(txtRqPath.Text);
                }
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtRqPath.Text = dialog.FileName;
                }
            }
        }
        
        private void BtnSave_Click(object? sender, EventArgs e)
        {
            string path = txtRqPath.Text.Trim();
            
            // Validate
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show(
                    "Please specify the location of rq.exe",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }
            
            if (!File.Exists(path))
            {
                MessageBox.Show(
                    "The specified file does not exist.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }
            
            if (!path.EndsWith("rq.exe", StringComparison.OrdinalIgnoreCase))
            {
                var result = MessageBox.Show(
                    "The selected file is not named 'rq.exe'. Continue anyway?",
                    "Warning",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (result != DialogResult.Yes)
                    return;
            }
            
            // Save
            var settings = SettingsManager.LoadSettings();
            settings.RqExePath = path;
            
            if (SettingsManager.SaveSettings(settings))
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }
}
```

**Add MenuStrip to Form1:**
- In Form1 designer, add MenuStrip
- Menu item: "Tools" → "Settings"
- Wire up to open SettingsForm

**Test:** 
- Open settings, select rq.exe path
- Test validation (empty path, non-existent file, wrong filename)
- Verify settings persist after app restart
- Check file created at `%APPDATA%\RqGui\settings.json`
- Test when AppData is not writable (shows warning but continues)

---

## Step 3A: Implement Simple Logger

**Action:** Create a basic logger for diagnostics and troubleshooting.

**Create `Logger.cs`:**
```csharp
using System;
using System.IO;

namespace RqGui
{
    public static class Logger
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RqGui",
            "rqgui.log"
        );
        
        private static readonly object lockObject = new object();
        
        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }
        
        public static void LogError(string message)
        {
            Log("ERROR", message);
        }
        
        public static void LogWarning(string message)
        {
            Log("WARN", message);
        }
        
        private static void Log(string level, string message)
        {
            try
            {
                lock (lockObject)
                {
                    string? directory = Path.GetDirectoryName(LogPath);
                    if (directory != null)
                    {
                        Directory.CreateDirectory(directory);
                    }
                    
                    string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";
                    File.AppendAllText(LogPath, logEntry + Environment.NewLine);
                }
            }
            catch
            {
                // Silently fail - logging should never crash the app
            }
        }
        
        public static void LogException(string context, Exception ex)
        {
            LogError($"{context}: {ex.GetType().Name} - {ex.Message}\n{ex.StackTrace}");
        }
    }
}
```

**Test:**
- Call Logger methods from various parts of the app
- Verify log file created at `%APPDATA%\RqGui\rqgui.log`
- Verify thread-safe (no corruption from concurrent writes)

---

## Step 4: Implement RQ Command Executor with Security

**Action:** Create service class to execute rq.exe safely without showing console window.

**Create `RqExecutor.cs`:**
```csharp
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RqGui
{
    public class SearchOptions
    {
        public bool UseGlob { get; set; }
        public bool CaseSensitive { get; set; }
        public bool IncludeHidden { get; set; }
        public string Extensions { get; set; } = "";
        public string FileType { get; set; } = "";
        public string SizeFilter { get; set; } = "";
        public int MaxResults { get; set; } = 1000;
    }
    
    public class SearchResult
    {
        public List<string> FilePaths { get; set; } = new List<string>();
        public string StandardError { get; set; } = "";
        public int ExitCode { get; set; }
        public bool WasCancelled { get; set; }
    }

    public class RqExecutor
    {
        private readonly string _rqPath;

        public RqExecutor(string rqPath)
        {
            _rqPath = rqPath;
        }

        /// <summary>
        /// Execute rq.exe search with proper argument escaping to prevent command injection.
        /// </summary>
        public async Task<SearchResult> ExecuteSearchAsync(
            string directory, 
            string searchTerm, 
            SearchOptions options,
            CancellationToken cancellationToken = default,
            IProgress<string>? progress = null)
        {
            var result = new SearchResult();
            
            // Thread-safe collection for results
            var results = new ConcurrentBag<string>();
            var errorOutput = new StringBuilder();
            
            try
            {
                // Validate directory exists
                if (!Directory.Exists(directory))
                {
                    throw new DirectoryNotFoundException($"Directory not found: {directory}");
                }
                
                // Build arguments list (SAFE - no command injection possible)
                var argumentsList = new List<string>
                {
                    directory,          // <where to look>
                    searchTerm          // <what to find>
                };
                
                // Add search options
                if (options.UseGlob)
                    argumentsList.Add("--glob");
                    
                if (options.CaseSensitive)
                    argumentsList.Add("--case");
                    
                if (options.IncludeHidden)
                    argumentsList.Add("--include-hidden");
                
                // Extension filter (comma-separated list)
                if (!string.IsNullOrWhiteSpace(options.Extensions))
                {
                    argumentsList.Add("--ext");
                    argumentsList.Add(options.Extensions.Trim());
                }
                
                // File type filter
                if (!string.IsNullOrWhiteSpace(options.FileType) && options.FileType.ToLower() != "any")
                {
                    argumentsList.Add("--type");
                    argumentsList.Add(options.FileType.ToLower());
                }
                
                // Size filter
                if (!string.IsNullOrWhiteSpace(options.SizeFilter))
                {
                    argumentsList.Add("--size");
                    argumentsList.Add(options.SizeFilter);
                }
                
                // Max results limit
                argumentsList.Add("--max-results");
                argumentsList.Add(options.MaxResults.ToString());
                
                Logger.LogInfo($"Executing rq with args: {string.Join(" ", argumentsList)}");
                
                var processInfo = new ProcessStartInfo
                {
                    FileName = _rqPath,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,              // CRITICAL: Hides console window
                    WindowStyle = ProcessWindowStyle.Hidden,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                
                // Add arguments safely (prevents command injection)
                foreach (var arg in argumentsList)
                {
                    processInfo.ArgumentList.Add(arg);
                }

                using (var process = new Process { StartInfo = processInfo })
                {
                    // Capture stdout
                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrWhiteSpace(e.Data))
                        {
                            results.Add(e.Data);
                            progress?.Report(e.Data);
                        }
                    };
                    
                    // Capture stderr
                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrWhiteSpace(e.Data))
                        {
                            errorOutput.AppendLine(e.Data);
                            Logger.LogWarning($"RQ stderr: {e.Data}");
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    
                    // Wait for completion or cancellation
                    await process.WaitForExitAsync(cancellationToken);
                    
                    result.ExitCode = process.ExitCode;
                    result.WasCancelled = cancellationToken.IsCancellationRequested;
                }
                
                result.FilePaths = results.ToList();
                result.StandardError = errorOutput.ToString();
                
                Logger.LogInfo($"Search completed. Found {result.FilePaths.Count} files. Exit code: {result.ExitCode}");
                
                if (!string.IsNullOrWhiteSpace(result.StandardError))
                {
                    Logger.LogWarning($"Search stderr output: {result.StandardError}");
                }
            }
            catch (OperationCanceledException)
            {
                result.WasCancelled = true;
                Logger.LogInfo("Search was cancelled by user");
            }
            catch (Exception ex)
            {
                Logger.LogException("Search execution failed", ex);
                throw;
            }
            
            return result;
        }

        public static bool ValidateRqPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;
            if (!File.Exists(path))
                return false;
            if (!path.EndsWith("rq.exe", StringComparison.OrdinalIgnoreCase))
                return false;
            return true;
        }
    }
}
```

**Key Security Features:**
- Uses `ArgumentList` property (available in .NET 5+) which properly escapes arguments
- No string concatenation or interpolation for command arguments
- Prevents command injection attacks
- Validates directory exists before executing
- Thread-safe result collection using `ConcurrentBag<string>`
- Captures both stdout and stderr
- Supports cancellation via `CancellationToken`
- Comprehensive logging for troubleshooting

**Test:**
- Create test method that runs simple search
- Verify no console window appears
- Verify output captured correctly
- Test with invalid rq.exe path (should handle gracefully)
- **Security Test**: Try malicious inputs like `"; del /Q *.*"` in search term - should be harmless
- Test cancellation works properly
- Verify stderr is captured and logged

---

## Step 5: Implement File Information Parser

**Action:** Parse rq output and extract file metadata.

**Create `FileResultItem.cs`:**
```csharp
using System;
using System.IO;

public class FileResultItem
{
    public string FileName { get; set; }
    public string FullPath { get; set; }
    public string Directory { get; set; }
    public long Size { get; set; }
    public DateTime Modified { get; set; }
    public string Extension { get; set; }

    public static FileResultItem FromPath(string path)
    {
        try
        {
            var fileInfo = new FileInfo(path);
            return new FileResultItem
            {
                FileName = fileInfo.Name,
                FullPath = fileInfo.FullName,
                Directory = fileInfo.DirectoryName,
                Size = fileInfo.Length,
                Modified = fileInfo.LastWriteTime,
                Extension = fileInfo.Extension
            };
        }
        catch
        {
            // If file no longer exists or permission denied
            return new FileResultItem
            {
                FileName = Path.GetFileName(path),
                FullPath = path,
                Directory = Path.GetDirectoryName(path),
                Size = 0,
                Modified = DateTime.MinValue,
                Extension = Path.GetExtension(path)
            };
        }
    }

    public string GetSizeFormatted()
    {
        string[] sizes = { "B", "KB", "MB", "GB", "TB" };
        double len = Size;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}
```

**Test:**
- Test with various file paths
- Verify size formatting works correctly
- Test with non-existent files (should not crash)

---

## Step 6: Populate ListView with Results and Cancellation

**Action:** Wire up search button to execute rq and display results with proper performance and cancellation.

**In `Form1.cs`, add fields and search handler:**
```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RqGui
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource? _searchCancellationSource;
        
        private async void btnSearch_Click(object sender, EventArgs e)
        {
            // Validate inputs
            var settings = SettingsManager.LoadSettings();
            if (!RqExecutor.ValidateRqPath(settings.RqExePath))
            {
                MessageBox.Show(
                    "Please configure rq.exe path in Tools → Settings",
                    "Configuration Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            if (string.IsNullOrWhiteSpace(txtSearchPath.Text))
            {
                MessageBox.Show(
                    "Please select a search directory",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                txtSearchPath.Focus();
                return;
            }
            
            if (string.IsNullOrWhiteSpace(txtSearchTerm.Text))
            {
                var result = MessageBox.Show(
                    "Search term is empty. This will list all files in the directory.\n\nContinue?",
                    "Confirm",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (result != DialogResult.Yes)
                    return;
            }

            // Prepare UI
            lvResults.Items.Clear();
            btnSearch.Enabled = false;
            btnCancel.Enabled = true;
            progressBar1.Visible = true;
            statusLabel.Text = "Searching...";
            
            // Create cancellation token
            _searchCancellationSource = new CancellationTokenSource();
            var cancellationToken = _searchCancellationSource.Token;

            try
            {
                // Build options
                var options = new SearchOptions
                {
                    UseGlob = chkGlob.Checked,
                    CaseSensitive = chkCaseSensitive.Checked,
                    IncludeHidden = chkIncludeHidden.Checked,
                    Extensions = txtExtensions.Text.Trim(),
                    FileType = cmbFileType.SelectedItem?.ToString() ?? "Any",
                    SizeFilter = BuildSizeFilter(),
                    MaxResults = (int)numMaxResults.Value
                };

                // Execute search
                var executor = new RqExecutor(settings.RqExePath);
                
                // Progress reporter (optional - updates status during search)
                var progress = new Progress<string>(filePath =>
                {
                    // Update status with last found file (don't add to list yet)
                    statusLabel.Text = $"Searching... Found: {Path.GetFileName(filePath)}";
                });
                
                var searchResult = await executor.ExecuteSearchAsync(
                    txtSearchPath.Text,
                    txtSearchTerm.Text,
                    options,
                    cancellationToken,
                    progress
                );

                // Check for cancellation
                if (searchResult.WasCancelled)
                {
                    statusLabel.Text = "Search cancelled";
                    Logger.LogInfo("User cancelled search");
                    return;
                }
                
                // Check for errors
                if (searchResult.ExitCode != 0 && !string.IsNullOrWhiteSpace(searchResult.StandardError))
                {
                    MessageBox.Show(
                        $"Search completed with errors:\n\n{searchResult.StandardError}",
                        "Search Errors",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }

                // Populate ListView with performance optimization
                lvResults.BeginUpdate();  // CRITICAL: Prevents flickering and improves performance
                try
                {
                    foreach (var filePath in searchResult.FilePaths)
                    {
                        var fileItem = FileResultItem.FromPath(filePath);
                        var listItem = new ListViewItem(new[]
                        {
                            fileItem.FileName,
                            fileItem.Directory ?? "",
                            fileItem.GetSizeFormatted(),
                            fileItem.Modified != DateTime.MinValue 
                                ? fileItem.Modified.ToString("yyyy-MM-dd HH:mm") 
                                : "N/A",
                            fileItem.Extension
                        })
                        {
                            Tag = fileItem // Store object for later use
                        };
                        lvResults.Items.Add(listItem);
                    }
                }
                finally
                {
                    lvResults.EndUpdate();  // Always call EndUpdate
                }

                statusLabel.Text = $"Found {searchResult.FilePaths.Count} files";
                
                if (searchResult.FilePaths.Count == options.MaxResults)
                {
                    statusLabel.Text += $" (limited to {options.MaxResults})";
                }
                
                // Save last search path
                settings.LastSearchPath = txtSearchPath.Text;
                SettingsManager.SaveSettings(settings);
            }
            catch (OperationCanceledException)
            {
                statusLabel.Text = "Search cancelled";
            }
            catch (Exception ex)
            {
                Logger.LogException("Search failed", ex);
                MessageBox.Show(
                    $"Search failed: {ex.Message}\n\nCheck logs for details.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                statusLabel.Text = "Search failed";
            }
            finally
            {
                progressBar1.Visible = false;
                btnSearch.Enabled = true;
                btnCancel.Enabled = false;
                _searchCancellationSource?.Dispose();
                _searchCancellationSource = null;
            }
        }
        
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (_searchCancellationSource != null && !_searchCancellationSource.IsCancellationRequested)
            {
                statusLabel.Text = "Cancelling search...";
                _searchCancellationSource.Cancel();
                btnCancel.Enabled = false;
            }
        }

        private string BuildSizeFilter()
        {
            if (cmbSizeOperator.SelectedIndex <= 0 || string.IsNullOrWhiteSpace(txtSizeValue.Text))
                return "";
            
            string op = cmbSizeOperator.SelectedItem?.ToString() ?? "";
            string value = txtSizeValue.Text.Trim();
            
            // RQ expects: +100M for "larger than" or -100M for "smaller than"
            return op == ">" ? $"+{value}" : $"-{value}";
        }
    }
}
```

**Key Improvements:**
- Uses `CancellationTokenSource` for proper cancellation support
- Implements Cancel button functionality
- Uses `BeginUpdate()`/`EndUpdate()` for ListView performance
- Validates all inputs before executing
- Handles errors from rq.exe gracefully
- Displays stderr if search had errors
- Saves last search path for convenience
- Progress reporting during search (optional)

**Test:**
- Run search with various filters
- Verify results appear in ListView quickly even with 1000+ files
- Test with no results (should show "Found 0 files")
- Test error handling with invalid directory
- **Test cancellation**: Start long search and click Cancel
- Verify UI remains responsive during large searches
- Test max results limit

---

## Step 7: Implement File and Folder Opening with Column Sorting

**Action:** Add double-click, context menu, and column sorting functionality.

**Add to `Form1.cs`:**
```csharp
private void InitializeListViewSorting()
{
    // Enable column header click to sort
    lvResults.ColumnClick += LvResults_ColumnClick;
    lvResults.ListViewItemSorter = new ListViewItemComparer(0, SortOrder.Ascending);
}

private void LvResults_ColumnClick(object? sender, ColumnClickEventArgs e)
{
    var currentSorter = lvResults.ListViewItemSorter as ListViewItemComparer;
    
    if (currentSorter != null && currentSorter.Column == e.Column)
    {
        // Same column - toggle sort order
        currentSorter.Order = currentSorter.Order == SortOrder.Ascending 
            ? SortOrder.Descending 
            : SortOrder.Ascending;
    }
    else
    {
        // New column - sort ascending
        lvResults.ListViewItemSorter = new ListViewItemComparer(e.Column, SortOrder.Ascending);
    }
    
    lvResults.Sort();
}

private void lvResults_DoubleClick(object sender, EventArgs e)
{
    if (lvResults.SelectedItems.Count == 0)
        return;

    var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
    if (fileItem != null)
    {
        OpenFile(fileItem.FullPath);
    }
}

private void OpenFile(string filePath)
{
    try
    {
        if (!File.Exists(filePath))
        {
            MessageBox.Show(
                "File no longer exists or cannot be accessed.",
                "File Not Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }
        
        Process.Start(new ProcessStartInfo
        {
            FileName = filePath,
            UseShellExecute = true
        });
        
        Logger.LogInfo($"Opened file: {filePath}");
    }
    catch (Exception ex)
    {
        Logger.LogException($"Failed to open file: {filePath}", ex);
        MessageBox.Show(
            $"Failed to open file: {ex.Message}",
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
}

private void OpenContainingFolder(string filePath)
{
    try
    {
        if (!File.Exists(filePath))
        {
            // Try to open the directory instead
            string? directory = Path.GetDirectoryName(filePath);
            if (directory != null && Directory.Exists(directory))
            {
                Process.Start("explorer.exe", $"\"{directory}\"");
                return;
            }
            
            MessageBox.Show(
                "File and directory no longer exist.",
                "Not Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }
        
        // Opens Explorer with file selected
        Process.Start("explorer.exe", $"/select,\"{filePath}\"");
        Logger.LogInfo($"Opened folder for: {filePath}");
    }
    catch (Exception ex)
    {
        Logger.LogException($"Failed to open folder: {filePath}", ex);
        MessageBox.Show(
            $"Failed to open folder: {ex.Message}",
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
}

private void CopyPathToClipboard(string path)
{
    try
    {
        Clipboard.SetText(path);
        statusLabel.Text = "Path copied to clipboard";
        Logger.LogInfo($"Copied to clipboard: {path}");
    }
    catch (Exception ex)
    {
        Logger.LogException("Failed to copy to clipboard", ex);
        MessageBox.Show(
            $"Failed to copy path: {ex.Message}",
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
}
```

**Add Context Menu:**
```csharp
private void InitializeContextMenu()
{
    var contextMenu = new ContextMenuStrip();
    
    var openItem = new ToolStripMenuItem("Open");
    openItem.Click += (s, e) =>
    {
        if (lvResults.SelectedItems.Count > 0)
        {
            var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
            if (fileItem != null)
                OpenFile(fileItem.FullPath);
        }
    };
    
    var openFolderItem = new ToolStripMenuItem("Open Containing Folder");
    openFolderItem.Click += (s, e) =>
    {
        if (lvResults.SelectedItems.Count > 0)
        {
            var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
            if (fileItem != null)
                OpenContainingFolder(fileItem.FullPath);
        }
    };
    
    var copyPathItem = new ToolStripMenuItem("Copy Full Path");
    copyPathItem.Click += (s, e) =>
    {
        if (lvResults.SelectedItems.Count > 0)
        {
            var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
            if (fileItem != null)
                CopyPathToClipboard(fileItem.FullPath);
        }
    };
    
    var copyFileNameItem = new ToolStripMenuItem("Copy File Name");
    copyFileNameItem.Click += (s, e) =>
    {
        if (lvResults.SelectedItems.Count > 0)
        {
            var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
            if (fileItem != null)
                CopyPathToClipboard(fileItem.FileName);
        }
    };
    
    contextMenu.Items.Add(openItem);
    contextMenu.Items.Add(openFolderItem);
    contextMenu.Items.Add(new ToolStripSeparator());
    contextMenu.Items.Add(copyPathItem);
    contextMenu.Items.Add(copyFileNameItem);
    
    lvResults.ContextMenuStrip = contextMenu;
}

// ListViewItemComparer for sorting
private class ListViewItemComparer : System.Collections.IComparer
{
    public int Column { get; set; }
    public SortOrder Order { get; set; }

    public ListViewItemComparer(int column, SortOrder order)
    {
        Column = column;
        Order = order;
    }

    public int Compare(object? x, object? y)
    {
        if (x == null || y == null)
            return 0;
            
        var itemX = (ListViewItem)x;
        var itemY = (ListViewItem)y;
        
        int result;
        
        // Special handling for size column (column 2)
        if (Column == 2)
        {
            var fileX = itemX.Tag as FileResultItem;
            var fileY = itemY.Tag as FileResultItem;
            
            if (fileX != null && fileY != null)
            {
                result = fileX.Size.CompareTo(fileY.Size);
            }
            else
            {
                result = string.Compare(itemX.SubItems[Column].Text, itemY.SubItems[Column].Text);
            }
        }
        // Special handling for date column (column 3)
        else if (Column == 3)
        {
            var fileX = itemX.Tag as FileResultItem;
            var fileY = itemY.Tag as FileResultItem;
            
            if (fileX != null && fileY != null)
            {
                result = fileX.Modified.CompareTo(fileY.Modified);
            }
            else
            {
                result = string.Compare(itemX.SubItems[Column].Text, itemY.SubItems[Column].Text);
            }
        }
        else
        {
            // Text comparison for other columns
            result = string.Compare(
                itemX.SubItems[Column].Text, 
                itemY.SubItems[Column].Text,
                StringComparison.OrdinalIgnoreCase
            );
        }
        
        return Order == SortOrder.Ascending ? result : -result;
    }
}

// Call in Form1 constructor
public Form1()
{
    InitializeComponent();
    InitializeContextMenu();
    InitializeListViewSorting();
    LoadInitialSettings();
}

private void LoadInitialSettings()
{
    var settings = SettingsManager.LoadSettings();
    if (!string.IsNullOrWhiteSpace(settings.LastSearchPath) && Directory.Exists(settings.LastSearchPath))
    {
        txtSearchPath.Text = settings.LastSearchPath;
    }
    
    Logger.LogInfo("Application started");
}
```

**Test:**
- Double-click a result (file should open in default application)
- Right-click → Open Containing Folder (Explorer opens with file selected)
- Right-click → Copy Full Path (paste in notepad to verify)
- Right-click → Copy File Name (paste to verify)
- Test with files that might not open (system files, etc.)
- Click column headers to sort (Name, Path, Size, Modified, Type)
- Click same column header again to reverse sort
- Verify size and date columns sort numerically/chronologically, not alphabetically

---

## Step 8: Add Folder Browser Dialog

**Action:** Implement directory selection.

**Add to `Form1.cs`:**
```csharp
private void btnBrowseDirectory_Click(object sender, EventArgs e)
{
    using (var dialog = new FolderBrowserDialog())
    {
        dialog.Description = "Select directory to search";
        dialog.ShowNewFolderButton = false;
        
        if (!string.IsNullOrWhiteSpace(txtSearchPath.Text))
        {
            dialog.SelectedPath = txtSearchPath.Text;
        }
        
        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtSearchPath.Text = dialog.SelectedPath;
        }
    }
}
```

**Test:**
- Click Browse button
- Navigate to various folders
- Verify selected path appears in textbox
- Test canceling dialog (path should remain unchanged)

---

## Step 9: Polish UI and Add Icons

**Action:** Enhance visual appearance to match Windows 11 style and add keyboard shortcuts.

**Apply Modern Styling:**
```csharp
// In Form1 Load event or constructor
private void ApplyModernStyling()
{
    // Form styling
    this.BackColor = Color.White;
    this.Font = new Font("Segoe UI", 9F);
    this.Text = "RQ File Search";
    this.MinimumSize = new Size(800, 600);
    
    // ListView styling
    lvResults.BackColor = Color.White;
    lvResults.BorderStyle = BorderStyle.FixedSingle;
    lvResults.Font = new Font("Segoe UI", 9F);
    
    // Button styling
    btnSearch.BackColor = Color.FromArgb(0, 120, 212); // Windows blue
    btnSearch.ForeColor = Color.White;
    btnSearch.FlatStyle = FlatStyle.Flat;
    btnSearch.FlatAppearance.BorderSize = 0;
    btnSearch.Cursor = Cursors.Hand;
    
    btnCancel.BackColor = Color.FromArgb(232, 17, 35); // Red
    btnCancel.ForeColor = Color.White;
    btnCancel.FlatStyle = FlatStyle.Flat;
    btnCancel.FlatAppearance.BorderSize = 0;
    btnCancel.Cursor = Cursors.Hand;
    
    btnBrowseDirectory.Cursor = Cursors.Hand;
    
    // Add placeholder text behavior
    SetPlaceholder(txtSearchTerm, "Enter search term or pattern");
    SetPlaceholder(txtExtensions, "e.g., jpg,png,pdf");
    SetPlaceholder(txtSizeValue, "e.g., 100M, 5K, 1G");
}

private void SetPlaceholder(TextBox textBox, string placeholder)
{
    // Simple placeholder implementation
    textBox.ForeColor = Color.Gray;
    textBox.Text = placeholder;
    
    textBox.GotFocus += (s, e) =>
    {
        if (textBox.Text == placeholder)
        {
            textBox.Text = "";
            textBox.ForeColor = Color.Black;
        }
    };
    
    textBox.LostFocus += (s, e) =>
    {
        if (string.IsNullOrWhiteSpace(textBox.Text))
        {
            textBox.Text = placeholder;
            textBox.ForeColor = Color.Gray;
        }
    };
}

private string GetTextBoxValue(TextBox textBox, string placeholder)
{
    return textBox.Text == placeholder || textBox.ForeColor == Color.Gray 
        ? "" 
        : textBox.Text;
}

private void Form1_KeyDown(object sender, KeyEventArgs e)
{
    // Ctrl+F to focus search
    if (e.Control && e.KeyCode == Keys.F)
    {
        txtSearchTerm.Focus();
        txtSearchTerm.SelectAll();
        e.Handled = true;
        e.SuppressKeyPress = true;
    }
    
    // Enter in search controls triggers search
    if (e.KeyCode == Keys.Enter && btnSearch.Enabled)
    {
        if (txtSearchTerm.Focused || txtSearchPath.Focused || 
            txtExtensions.Focused || txtSizeValue.Focused)
        {
            btnSearch.PerformClick();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
    }
    
    // Escape cancels search
    if (e.KeyCode == Keys.Escape && btnCancel.Enabled)
    {
        btnCancel.PerformClick();
        e.Handled = true;
        e.SuppressKeyPress = true;
    }
    
    // Ctrl+O opens selected file
    if (e.Control && e.KeyCode == Keys.O && lvResults.SelectedItems.Count > 0)
    {
        var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
        if (fileItem != null)
            OpenFile(fileItem.FullPath);
        e.Handled = true;
        e.SuppressKeyPress = true;
    }
    
    // Ctrl+Shift+O opens containing folder
    if (e.Control && e.Shift && e.KeyCode == Keys.O && lvResults.SelectedItems.Count > 0)
    {
        var fileItem = lvResults.SelectedItems[0].Tag as FileResultItem;
        if (fileItem != null)
            OpenContainingFolder(fileItem.FullPath);
        e.Handled = true;
        e.SuppressKeyPress = true;
    }
}

// Update Form1 constructor
public Form1()
{
    InitializeComponent();
    this.KeyPreview = true;  // Enable form to receive key events first
    this.KeyDown += Form1_KeyDown;
    
    InitializeContextMenu();
    InitializeListViewSorting();
    ApplyModernStyling();
    LoadInitialSettings();
}
```

**Update search handler to use placeholder-aware method:**
```csharp
// In btnSearch_Click, replace direct text access with:
var searchTerm = GetTextBoxValue(txtSearchTerm, "Enter search term or pattern");
var extensions = GetTextBoxValue(txtExtensions, "e.g., jpg,png,pdf");
var sizeValue = GetTextBoxValue(txtSizeValue, "e.g., 100M, 5K, 1G");

// Then use these variables in SearchOptions
var options = new SearchOptions
{
    UseGlob = chkGlob.Checked,
    CaseSensitive = chkCaseSensitive.Checked,
    IncludeHidden = chkIncludeHidden.Checked,
    Extensions = extensions,
    FileType = cmbFileType.SelectedItem?.ToString() ?? "Any",
    SizeFilter = BuildSizeFilter(), // Update this method too
    MaxResults = (int)numMaxResults.Value
};
```

**Update BuildSizeFilter:**
```csharp
private string BuildSizeFilter()
{
    string value = GetTextBoxValue(txtSizeValue, "e.g., 100M, 5K, 1G");
    
    if (cmbSizeOperator.SelectedIndex <= 0 || string.IsNullOrWhiteSpace(value))
        return "";
    
    string op = cmbSizeOperator.SelectedItem?.ToString() ?? "";
    
    // RQ expects: +100M for "larger than" or -100M for "smaller than"
    return op == ">" ? $"+{value}" : $"-{value}";
}
```

**Test:**
- Verify modern appearance
- Test keyboard shortcuts:
  - Ctrl+F (focus search)
  - Enter (start search)
  - Escape (cancel search)
  - Ctrl+O (open file)
  - Ctrl+Shift+O (open folder)
- Check contrast and readability
- Verify placeholder text appears and disappears correctly
- Test tab order flows logically

---

## Step 10: Add Progress Indication (Already Integrated)

**Note:** Progress indication has already been integrated in Step 6 with:
- `progressBar1` shown/hidden during search
- Status label updates with progress
- Cancel button enabled during search

**Additional Enhancement - Add File Count Updates:**

Already implemented in Step 6 with the Progress reporter:
```csharp
var progress = new Progress<string>(filePath =>
{
    statusLabel.Text = $"Searching... Found: {Path.GetFileName(filePath)}";
});
```

**Test:**
- Run search on large directory
- Verify progress bar animates during search
- Verify status updates show found files
- Verify UI remains responsive (can move window, click Cancel, etc.)
- Test Cancel button interrupts search immediately

---

## Step 11: Error Handling and Edge Cases (Enhanced)

**Action:** Handle common error scenarios gracefully. Most error handling has been integrated into previous steps, but add Form_Load validation.

**Add comprehensive Form_Load validation:**
```csharp
private void Form1_Load(object sender, EventArgs e)
{
    try
    {
        string? rqPath = GetRqPath();
        if (rqPath == null)
        {
            var result = MessageBox.Show(
                "rq.exe not found on your system.\n\n" +
                "Would you like to configure it now?\n\n" +
                "(You can also configure it later via Tools → Settings)",
                "RQ Configuration Required",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            
            if (result == DialogResult.Yes)
            {
                using (var settingsForm = new SettingsForm())
                {
                    settingsForm.ShowDialog(this);
                }
            }
            else
            {
                statusLabel.Text = "RQ.exe not configured - configure via Tools → Settings";
            }
        }
        else
        {
            statusLabel.Text = "Ready";
            Logger.LogInfo($"Using rq.exe at: {rqPath}");
        }
    }
    catch (Exception ex)
    {
        Logger.LogException("Form load error", ex);
        MessageBox.Show(
            $"Error during startup: {ex.Message}",
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
}

// Check if rq.exe is on PATH or configured
private string? GetRqPath()
{
    var settings = SettingsManager.LoadSettings();
    
    // First check configured path
    if (RqExecutor.ValidateRqPath(settings.RqExePath))
        return settings.RqExePath;
    
    // Try PATH
    string? pathEnv = Environment.GetEnvironmentVariable("PATH");
    if (pathEnv != null)
    {
        string[] pathDirs = pathEnv.Split(Path.PathSeparator);
        foreach (var dir in pathDirs)
        {
            try
            {
                string testPath = Path.Combine(dir.Trim(), "rq.exe");
                if (File.Exists(testPath))
                {
                    // Found it - save for future use
                    settings.RqExePath = testPath;
                    SettingsManager.SaveSettings(settings);
                    return testPath;
                }
            }
            catch
            {
                // Skip invalid paths in PATH variable
                continue;
            }
        }
    }
    
    return null;
}
```

**Add protection against rapid repeated clicks:**
```csharp
private async void btnSearch_Click(object sender, EventArgs e)
{
    // Prevent multiple simultaneous searches
    if (!btnSearch.Enabled)
        return;
        
    // ... rest of existing code ...
}
```

**Test Edge Cases:**
- [x] Launch app without rq.exe configured (already tested in Step 3)
- [x] Search with special characters in path (protected by ArgumentList)
- [ ] Search directory that user doesn't have permission to access
- [ ] Search with empty search term (already prompts in Step 6)
- [ ] Very large result sets (1000+ files) - handled by BeginUpdate/EndUpdate
- [ ] Network drives and UNC paths
- [ ] Cancel search mid-operation (already implemented in Step 6)
- [ ] Files deleted between search and opening (already checked in Step 7)
- [ ] Rapid button clicking (now prevented)
- [ ] Invalid size values (e.g., "abc", "100XYZ")
- [ ] Unicode filenames and paths

---

## Step 12: Comprehensive Testing

**Action:** Systematic testing of all functionality.

**Test Checklist:**

**Basic Functionality:**
- [ ] App launches without errors
- [ ] Settings save and load correctly
- [ ] Settings form validates inputs properly
- [ ] Directory browser works
- [ ] Search executes and returns results
- [ ] Results display in ListView with correct data
- [ ] Double-click opens files
- [ ] Right-click menu works (open, open folder, copy path, copy name)
- [ ] Cancellation works mid-search

**Filter Testing:**
- [ ] Glob patterns work (*.txt, file?.doc, *.{jpg,png}, etc.)
- [ ] Case sensitive search works
- [ ] Include hidden files works
- [ ] Extension filter works (single: jpg, multiple: jpg,png,pdf)
- [ ] File type filter works (image, video, audio, document)
- [ ] Size filter works (>100M, <1K, +5G, -500K)
- [ ] Max results limit works
- [ ] Empty search term behavior (lists all files with confirmation)

**UI/UX Testing:**
- [ ] No console window appears during search
- [ ] Progress indication shows during search
- [ ] Status bar updates correctly with file count
- [ ] Keyboard shortcuts work (Ctrl+F, Enter, Escape, Ctrl+O, Ctrl+Shift+O)
- [ ] UI remains responsive during search
- [ ] Column sorting works (all 5 columns)
- [ ] Column sorting toggles ascending/descending
- [ ] Size and date columns sort numerically/chronologically
- [ ] Window resizes properly
- [ ] Tab order flows logically
- [ ] Placeholder text works in textboxes
- [ ] Cancel button only enabled during search

**Error Handling:**
- [ ] Invalid rq.exe path handled gracefully
- [ ] Non-existent directory handled (shows error before search)
- [ ] Permission denied on directory handled
- [ ] Empty results handled (shows "Found 0 files")
- [ ] Invalid file paths in results handled
- [ ] File deleted after search but before open handled
- [ ] rq.exe stderr output displayed to user
- [ ] Rapid button clicking doesn't cause issues
- [ ] Invalid size filter values handled

**Security Testing:**
- [ ] Command injection attempts fail safely (try search terms with quotes, semicolons, ampersands)
- [ ] Special characters in paths handled (Unicode, spaces, etc.)
- [ ] UNC paths work correctly

**Performance:**
- [ ] Search on large directory (10,000+ files)
- [ ] ListView population is fast (BeginUpdate/EndUpdate working)
- [ ] UI doesn't freeze during search
- [ ] Memory usage reasonable
- [ ] Search can be cancelled quickly

**Logging:**
- [ ] Log file created at %APPDATA%\RqGui\rqgui.log
- [ ] Errors logged with stack traces
- [ ] Info messages logged for key actions
- [ ] Thread-safe (no log corruption)

**Edge Cases:**
- [ ] Network drives that become unavailable
- [ ] Paths with spaces
- [ ] Unicode filenames
- [ ] Very long file paths (near MAX_PATH)
- [ ] Files with no extension
- [ ] Read-only file systems
- [ ] Multiple app instances (don't corrupt settings)

---

## Step 13: Documentation

**Action:** Create comprehensive README and usage guide.

**Create `README.md`:**
```markdown
# RQ GUI

A modern Windows GUI application for the `rq` command-line file search utility. Provides a native Windows File Explorer-like interface for fast and powerful file searching.

## Features

- **Fast File Search**: Leverages the rq search engine for rapid file discovery
- **Flexible Filtering**: 
  - Glob patterns (wildcards: *, ?, [], {})
  - File extension filtering
  - File type filtering (image, video, audio, document)
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
- **Safe and Secure**: Command injection protection
- **Performance**: Handles 10,000+ results efficiently

## Requirements

- **Windows 11** (or Windows 10 with .NET 6.0 Runtime)
- **.NET 6.0 Runtime** or later ([Download](https://dotnet.microsoft.com/download/dotnet/6.0))
- **rq.exe** - Download from [github.com/seeyebe/rq](https://github.com/seeyebe/rq/releases)

## Installation

### Option 1: Download Release Binary
1. Download the latest release from the [Releases page](#)
2. Extract `RqGui.exe` to a folder of your choice
3. Download `rq.exe` from [rq releases](https://github.com/seeyebe/rq/releases)
4. Run `RqGui.exe`
5. Go to **Tools → Settings** and point to your `rq.exe` location

### Option 2: Build from Source
1. Clone this repository
2. Open in Visual Studio Code or Visual Studio
3. Run:
   ```powershell
   dotnet build -c Release
   ```
4. Executable will be in `bin/Release/net6.0-windows/`

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
- Select from dropdown: image, video, audio, document, or Any

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
| `Ctrl+C` | Copy selected file path (via right-click menu) |

### Context Menu (Right-Click)
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

You can manually edit this file or use **Tools → Settings** menu.

## Logs

Application logs are written to: `%APPDATA%\RqGui\rqgui.log`

If you encounter issues, check this file for error details.

## Troubleshooting

**"rq.exe not found"**
- Configure rq.exe location via Tools → Settings
- Or place rq.exe in your PATH environment variable

**No results found**
- Check directory path is correct
- Try without filters first
- Verify search term spelling
- Check case-sensitive option if enabled

**Search is slow**
- Use more specific search terms
- Enable glob patterns only when needed
- Reduce search depth with max-depth (future feature)
- Use extension/type filters to narrow search

**Files won't open**
- Verify file still exists (may have been deleted/moved)
- Check file associations in Windows
- Try "Open Containing Folder" instead

**Permission errors**
- Run as Administrator for system directories
- Some folders require elevated privileges

## Known Limitations

- Maximum 100,000 results per search (adjustable via Max Results setting)
- Long file paths (>260 characters) may have issues on older Windows versions
- Network drives may be slow or require specific permissions

## License

[Your License Here]

## Credits

- Built with .NET 6.0 WinForms
- Uses [rq](https://github.com/seeyebe/rq) by seeyebe for file searching
```

**Add XML documentation comments to all public classes and methods:**

Example for RqExecutor.cs:
```csharp
/// <summary>
/// Executes rq.exe file search utility and captures results.
/// </summary>
public class RqExecutor
{
    /// <summary>
    /// Executes a search using rq.exe with the specified parameters.
    /// </summary>
    /// <param name="directory">The root directory to search in.</param>
    /// <param name="searchTerm">The search pattern or term to find.</param>
    /// <param name="options">Search options and filters.</param>
    /// <param name="cancellationToken">Token to cancel the operation.</param>
    /// <param name="progress">Progress reporter for real-time updates.</param>
    /// <returns>A SearchResult containing found file paths and execution details.</returns>
    /// <exception cref="DirectoryNotFoundException">Thrown when the directory does not exist.</exception>
    public async Task<SearchResult> ExecuteSearchAsync(...)
    {
        // implementation
    }
}
```

**Test Documentation:**
- Verify README is clear and complete
- Test all examples in README
- Verify links work
- Check that troubleshooting steps are accurate

---

## Step 14: Build and Package

**Action:** Create release build with proper configuration.

**Build release:**
```powershell
# Self-contained (includes .NET runtime - larger file)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

# Framework-dependent (requires .NET 6.0 installed - smaller file)
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
```

**Output location:**
```
bin/Release/net6.0-windows/win-x64/publish/RqGui.exe
```

**Create release package:**
```powershell
# Create release folder
New-Item -ItemType Directory -Path "release" -Force

# Copy executable
Copy-Item "bin/Release/net6.0-windows/win-x64/publish/RqGui.exe" -Destination "release/"

# Copy README
Copy-Item "README.md" -Destination "release/"

# Create ZIP package
Compress-Archive -Path "release/*" -DestinationPath "RqGui-v1.0.0-win-x64.zip" -Force
```

**Test published executable:**
- Extract ZIP to clean folder
- Run RqGui.exe
- Verify it works without Visual Studio
- Test on clean Windows 11 installation (if possible)
- Test with/without .NET 6.0 Runtime (depending on build type)
- Verify icon appears correctly
- Verify all features work from published build

**Optional - Code Signing (for production):**
```powershell
# Sign the executable with a code signing certificate
signtool sign /f "YourCertificate.pfx" /p "password" /t http://timestamp.digicert.com RqGui.exe
```

---

## Step 15: Final Verification

**Action:** Complete end-to-end testing before release.

**Full Workflow Test:**
1. Launch app fresh (delete settings file first)
2. Verify rq.exe detection or configuration prompt
3. Configure rq.exe path via Settings
4. Select search directory
5. Perform basic search
6. Apply various filters and repeat searches
7. Test column sorting
8. Open files and folders from results
9. Test keyboard shortcuts
10. Cancel a long-running search
11. Test context menu options
12. Change settings again
13. Re-launch app and verify settings persist
14. Test all menu items
15. Trigger intentional errors and verify graceful handling
16. Check log file for proper entries

**Sign-off Criteria:**
- [ ] All features work as specified
- [ ] No console window appears during any operation
- [ ] UI is responsive and modern (Windows 11 style)
- [ ] Error handling is robust (no unhandled exceptions)
- [ ] Settings persist across sessions
- [ ] Documentation is complete and accurate
- [ ] Logging works and is thread-safe
- [ ] Security: No command injection possible
- [ ] Performance: 10,000+ results load quickly
- [ ] All keyboard shortcuts functional
- [ ] Column sorting works correctly
- [ ] Cancellation works reliably
- [ ] Published build works independently
- [ ] Application icon displays correctly

**Final Code Review:**
- [ ] All classes have XML documentation
- [ ] No hardcoded paths (except %APPDATA%)
- [ ] All using statements are necessary
- [ ] No compiler warnings
- [ ] Nullable reference types handled correctly
- [ ] Dispose pattern used for IDisposable objects
- [ ] No race conditions or threading issues
- [ ] Logging at appropriate levels (Info, Warning, Error)

---

## Troubleshooting

**If console window appears:**
- Verify `CreateNoWindow = true` in ProcessStartInfo
- Verify `UseShellExecute = false`
- Check if rq.exe is spawning child processes
- Ensure `WindowStyle = ProcessWindowStyle.Hidden` is set

**If UI freezes during search:**
- Ensure `async/await` is used correctly in btnSearch_Click
- Verify `ExecuteSearchAsync` properly uses `Task` and `await`
- Check that `BeginUpdate/EndUpdate` wraps ListView population
- Ensure progress updates happen on UI thread via `Progress<T>`
- Verify `CancellationToken` is properly threaded through

**If files won't open:**
- Verify `UseShellExecute = true` when opening files (in OpenFile method)
- Check file associations in Windows (Control Panel → Default Programs)
- Test with known file types (txt, jpg, pdf)
- Check if file was deleted/moved after search
- Try "Open Containing Folder" to verify file location

**If ListView is slow with many results:**
- Confirm `BeginUpdate()` is called before populating
- Confirm `EndUpdate()` is called in finally block
- Check if event handlers are firing excessively
- Consider reducing max results limit

**If search returns no results but files exist:**
- Verify search directory path is correct
- Check case-sensitive option (try with it off)
- Try glob pattern if searching for extensions (*.txt)
- Check if rq.exe has permission to access directory
- Review rq.exe stderr output for errors
- Test rq.exe directly from command line with same parameters

**If cancellation doesn't work:**
- Verify `CancellationTokenSource` is created before search
- Check that token is passed to `ExecuteSearchAsync`
- Ensure `process.WaitForExitAsync(cancellationToken)` receives token
- Verify btnCancel calls `_searchCancellationSource.Cancel()`

**If settings don't persist:**
- Check write permissions to %APPDATA%\RqGui folder
- Verify `SettingsManager.SaveSettings()` is being called
- Check for exceptions in log file
- Try running as Administrator
- Check if antivirus is blocking file writes

**If command injection seems possible:**
- Verify using `ArgumentList` property, NOT string concatenation
- Never use `Arguments = "string"` property
- Check that all user input goes through `ArgumentList.Add()`
- Review RqExecutor.cs implementation

**If log file has issues:**
- Check write permissions to %APPDATA%\RqGui folder
- Verify lock object is used consistently
- Check if multiple app instances are running
- Verify Directory.CreateDirectory succeeds

**If keyboard shortcuts don't work:**
- Ensure `Form.KeyPreview = true` is set
- Verify `KeyDown` event handler is wired up
- Check that `e.Handled = true` and `e.SuppressKeyPress = true` are set
- Ensure no controls are consuming events first

**If column sorting doesn't work:**
- Verify `ListView.Sorting` property is NOT set to None
- Check that `ListViewItemSorter` is assigned
- Ensure `ColumnClick` event is wired up
- Verify `ListView.Sort()` is called after changing sorter

**Performance Issues:**

*Large result sets (10,000+ files):*
- Use `BeginUpdate/EndUpdate` (mandatory)
- Increase max results limit cautiously
- Consider enabling rq.exe's `--max-results` flag
- Check memory usage in Task Manager

*Slow search execution:*
- Not a GUI issue - rq.exe performance
- Try more specific search terms
- Use extension filters to narrow scope
- Check disk speed (especially network drives)
- Enable `--stats` flag in rq to diagnose

**Debugging Tips:**

1. **Check the log file first:** `%APPDATA%\RqGui\rqgui.log`
2. **Test rq.exe independently:** Run from PowerShell to isolate GUI vs rq issues
3. **Enable verbose logging:** Add more Logger.LogInfo() calls temporarily
4. **Use Process Explorer:** Check if rq.exe process is actually running
5. **Test with minimal filters:** Start with simplest search and add filters incrementally

**Common User Errors:**

- Forgetting to check "Use Glob Patterns" when using wildcards
- Entering `100M` without selecting `>` or `<` operator
- Using regex patterns without rq regex support (use glob instead)
- Searching system folders without admin rights
- Expecting instant results on network drives

---

## Future Enhancements (Post-V1)

**Priority: High**
- **Date Filters UI**: Add date range pickers for --after and --before flags
- **Max Depth Control**: Add numeric control for --max-depth to limit recursion
- **Search History**: Dropdown of recent searches for quick re-execution
- **File Preview Panel**: Show file contents for text files in a preview pane
- **Multiple Select Operations**: Allow selecting multiple files for batch operations
- **Export Results**: Export search results to CSV or Excel

**Priority: Medium**
- **Saved Search Presets**: Save frequently used filter combinations
- **Dark Mode Theme**: Toggle between light and dark UI themes
- **Regex Support**: UI option for --regex flag with syntax highlighting
- **Follow Symlinks Option**: Checkbox for --follow-symlinks flag
- **No-Skip Option**: Checkbox for --no-skip to search all directories
- **Real-time Stats**: Display --stats output (files scanned, speed, etc.)
- **Thread Count Control**: Allow user to specify --threads parameter
- **Timeout Setting**: Configure --timeout for automatic search cancellation

**Priority: Low**
- **Recent Directories**: Quick access to recently searched directories
- **Favorites/Bookmarks**: Save frequently searched locations
- **Custom File Type Definitions**: Let users define custom file type groups
- **Portable Mode**: Option to store settings in app directory vs AppData
- **Multi-language Support**: Internationalization (i18n)
- **Accessibility Improvements**: Screen reader support, high contrast mode
- **File Icon Display**: Show appropriate icons for file types in ListView
- **Drag & Drop**: Drag folder into app to set search directory
- **Integration with Everything**: Fallback to Everything search engine if rq.exe unavailable
- **PowerToys Integration**: Become a PowerToys module
- **JSON Output Mode**: Use --json flag for structured output parsing

**User-Requested Features** (track after v1 release):
- Watch mode (continuous monitoring)
- Find duplicate files
- Content search (grep-like functionality)
- Batch rename from results
- File comparison tools
- Integration with version control (ignore .git, etc.)

**Technical Debt & Improvements:**
- Unit tests for core logic
- Integration tests for UI
- CI/CD pipeline setup
- Automated installer creation
- Code signing certificate
- Windows Store distribution
- Auto-update mechanism
- Telemetry and crash reporting (opt-in)
- Performance profiling and optimization
- Memory usage optimization for very large result sets
- Virtual mode ListView for 100,000+ results

---

## Appendix: RQ Command Reference

**From:** https://github.com/seeyebe/rq

### Basic Syntax
```
rq <where to look> <what to find> [options]
```

### Common Options

| Flag | Description | Example |
|------|-------------|---------|
| `--glob` | Enable wildcard patterns | `rq C:\ "*.txt" --glob` |
| `--regex` | Use regex patterns | `rq . "file[0-9]+\.txt" --regex` |
| `--case` | Case-sensitive search | `rq . "README" --case` |
| `--include-hidden` | Include hidden files | `rq . config --include-hidden` |
| `--ext <exts>` | Filter by extensions | `rq . report --ext pdf,docx` |
| `--type <type>` | Filter by file type | `rq . "" --type image` |
| `--size +<size>` | Files larger than size | `rq D:\ "" --size +100M` |
| `--size -<size>` | Files smaller than size | `rq C:\ "" --size -1K` |
| `--after <date>` | Modified after date | `rq . doc --after 2025-01-01` |
| `--before <date>` | Modified before date | `rq . tmp --before 2024-01-01` |
| `--max-depth <n>` | Max directory depth | `rq . file --max-depth 2` |
| `--max-results <n>` | Stop after n results | `rq C:\ "" --max-results 100` |
| `--follow-symlinks` | Follow symbolic links | `rq . link --follow-symlinks` |
| `--no-skip` | Don't skip special dirs | `rq . node --no-skip` |
| `--json` | JSON output format | `rq . file --json` |
| `--preview` | Show file previews | `rq . script --preview` |
| `--stats` | Show search statistics | `rq D:\ backup --stats` |
| `--threads <n>` | Number of threads | `rq C:\ file --threads 12` |
| `--timeout <ms>` | Timeout in milliseconds | `rq C:\ "" --timeout 30000` |

### Built-in File Types

| Type | Extensions |
|------|-----------|
| `image` | jpg, jpeg, png, gif, bmp, svg, webp, ico |
| `video` | mp4, avi, mkv, mov, wmv, flv, webm, m4v |
| `audio` | mp3, wav, flac, aac, ogg, wma, m4a |
| `document` | pdf, doc, docx, xls, xlsx, ppt, pptx, txt, md |
| `archive` | zip, rar, 7z, tar, gz, bz2, xz |
| `code` | c, cpp, h, hpp, cs, java, py, js, ts, go, rs |

### Size Units
- `K` or `KB` - Kilobytes
- `M` or `MB` - Megabytes
- `G` or `GB` - Gigabytes
- `T` or `TB` - Terabytes

### Glob Patterns
- `*` - Matches any characters
- `?` - Matches single character
- `[abc]` - Matches one character: a, b, or c
- `[a-z]` - Matches one character in range
- `{jpg,png}` - Matches jpg OR png

### Examples
```powershell
# Find all C files
rq C:\Dev "*.c" --glob

# Find large files
rq C:\ "" --size +100M

# Recent documents
rq C:\Users\me\Documents report --ext pdf,docx --after 2025-01-01

# Images in current directory only
rq . "" --type image --max-depth 1

# Case-sensitive search for README
rq . README --case

# Search with statistics
rq D:\ "backup" --stats --threads 12
```