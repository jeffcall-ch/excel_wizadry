using System.Diagnostics;

namespace RqGui;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
        Logger.LogInfo("Application started");

        var settings = SettingsManager.LoadSettings();
        if (!string.IsNullOrEmpty(settings.LastSearchPath))
        {
            txtSearchPath.Text = settings.LastSearchPath;
        }

        InitializeContextMenu();
        InitializeListViewSorting();
    }

    private void settingsToolStripMenuItem_Click(object? sender, EventArgs e)
    {
        using var settingsForm = new SettingsForm();
        settingsForm.ShowDialog();
    }

    private void btnBrowse_Click(object? sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog();
        dialog.Description = "Select a folder to search";
        
        if (!string.IsNullOrEmpty(txtSearchPath.Text) && Directory.Exists(txtSearchPath.Text))
        {
            dialog.SelectedPath = txtSearchPath.Text;
        }

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtSearchPath.Text = dialog.SelectedPath;
        }
    }

    private async void btnSearch_Click(object? sender, EventArgs e)
    {
        string searchPath = txtSearchPath.Text;
        string searchTerm = txtSearchTerm.Text;

        if (string.IsNullOrWhiteSpace(searchPath))
        {
            MessageBox.Show("Please enter a directory to search.", "Missing Directory", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!Directory.Exists(searchPath))
        {
            MessageBox.Show($"Directory not found: {searchPath}", "Invalid Directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Get rq.exe path - will use embedded version if no custom path configured
        string rqExePath;
        try
        {
            rqExePath = EmbeddedResourceHelper.GetRqExePath();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to locate rq.exe: {ex.Message}", "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var settings = SettingsManager.LoadSettings();
        settings.LastSearchPath = searchPath;
        SettingsManager.SaveSettings(settings);

        var options = new SearchOptions
        {
            UseGlob = chkGlob.Checked,
            UseRegex = chkRegex.Checked,
            CaseSensitive = chkCaseSensitive.Checked,
            IncludeHidden = chkIncludeHidden.Checked,
            Extensions = chkExtensions.Checked ? txtExtensions.Text : "",
            FileType = chkFileType.Checked ? txtFileType.Text : "",
            SizeFilter = chkSize.Checked ? NormalizeSizeFilter(txtSize.Text) : "",
            DateAfter = chkDateAfter.Checked ? dtpDateAfter.Value.ToString("yyyy-MM-dd") : "",
            DateBefore = chkDateBefore.Checked ? dtpDateBefore.Value.ToString("yyyy-MM-dd") : "",
            MaxDepth = chkMaxDepth.Checked && int.TryParse(txtMaxDepth.Text, out int depth) ? depth : 0,
            Threads = chkThreads.Checked && int.TryParse(txtThreads.Text, out int threads) ? threads : 0,
            Timeout = chkTimeout.Checked && int.TryParse(txtTimeout.Text, out int timeout) ? timeout : 0
        };

        btnSearch.Enabled = false;
        statusLabel.Text = "Searching...";
        lvResults.Items.Clear();

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            var executor = new RqExecutor(rqExePath);
            var result = await executor.ExecuteSearchAsync(searchPath, searchTerm, options, CancellationToken.None);

            if (result.ExitCode == 0)
            {
                foreach (var filePath in result.FilePaths)
                {
                    if (File.Exists(filePath))
                    {
                        var item = FileResultItem.FromPath(filePath);
                        var listItem = new ListViewItem(item.FileName);
                        listItem.SubItems.Add(item.FullPath);
                        listItem.SubItems.Add(item.GetSizeFormatted());
                        listItem.SubItems.Add(item.Modified.ToString("yyyy-MM-dd HH:mm:ss"));
                        listItem.SubItems.Add(item.Created.ToString("yyyy-MM-dd HH:mm:ss"));
                        listItem.Tag = item; // Store the item for context menu and double-click
                        lvResults.Items.Add(listItem);
                    }
                }

                stopwatch.Stop();

                // Auto-resize columns except Path (keep Path at fixed width)
                lvResults.Columns[0].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent); // Name
                lvResults.Columns[2].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent); // Size
                lvResults.Columns[3].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent); // Modified
                lvResults.Columns[4].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent); // Created

                statusLabel.Text = $"Found {result.FilePaths.Count} file(s) in {stopwatch.ElapsedMilliseconds} ms";
                Logger.LogInfo($"Search completed: {result.FilePaths.Count} files found in {stopwatch.ElapsedMilliseconds} ms");
            }
            else
            {
                stopwatch.Stop();
                statusLabel.Text = "Search failed";
                MessageBox.Show($"Search failed with exit code {result.ExitCode}", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        catch (Exception ex)
        {
            Logger.LogException("Search error", ex);
            statusLabel.Text = "Search error";
            MessageBox.Show($"Error during search: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnSearch.Enabled = true;
        }
    }

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

    private void lvResults_DoubleClick(object? sender, EventArgs e)
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
                    Process.Start("explorer.exe", directory);
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

    private void InitializeListViewSorting()
    {
        // Enable column header click to sort
        lvResults.ColumnClick += LvResults_ColumnClick;
        lvResults.ListViewItemSorter = new ListViewItemComparer(0, SortOrder.Ascending);
    }

    private void LvResults_ColumnClick(object? sender, ColumnClickEventArgs e)
    {
        var currentSorter = lvResults.ListViewItemSorter as ListViewItemComparer;

        // Clear all column header arrows first
        for (int i = 0; i < lvResults.Columns.Count; i++)
        {
            string headerText = lvResults.Columns[i].Text;
            // Remove existing arrows
            headerText = headerText.Replace(" ▲", "").Replace(" ▼", "");
            lvResults.Columns[i].Text = headerText;
        }

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
            currentSorter = lvResults.ListViewItemSorter as ListViewItemComparer;
        }

        // Add arrow to current sorted column
        if (currentSorter != null)
        {
            string arrow = currentSorter.Order == SortOrder.Ascending ? " ▲" : " ▼";
            string headerText = lvResults.Columns[e.Column].Text;
            if (!headerText.Contains("▲") && !headerText.Contains("▼"))
            {
                lvResults.Columns[e.Column].Text = headerText + arrow;
            }
        }

        lvResults.Sort();
    }

    /// <summary>
    /// Normalize size filter to rq.exe format.
    /// Converts common symbols like > to + and < to -.
    /// Examples: ">1MB" -> "+1MB", "<500K" -> "-500K"
    /// </summary>
    private string NormalizeSizeFilter(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return "";

        string normalized = input.Trim();

        // Replace > with + and < with -
        if (normalized.StartsWith(">"))
            normalized = "+" + normalized.Substring(1).TrimStart();
        else if (normalized.StartsWith("<"))
            normalized = "-" + normalized.Substring(1).TrimStart();

        // If no prefix, assume greater than
        if (!normalized.StartsWith("+") && !normalized.StartsWith("-"))
        {
            // Check if it starts with a digit
            if (normalized.Length > 0 && char.IsDigit(normalized[0]))
                normalized = "+" + normalized;
        }

        return normalized;
    }
}

/// <summary>
/// Compares ListView items for sorting by different columns.
/// </summary>
class ListViewItemComparer : System.Collections.IComparer
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

        int result = 0;

        // Column 0: Name (string)
        // Column 1: Path (string)
        // Column 2: Size (need to compare actual file size, not formatted string)
        // Column 3: Modified (datetime)
        // Column 4: Created (datetime)

        if (Column == 2) // Size column
        {
            var fileX = itemX.Tag as FileResultItem;
            var fileY = itemY.Tag as FileResultItem;
            if (fileX != null && fileY != null)
            {
                result = fileX.Size.CompareTo(fileY.Size);
            }
        }
        else if (Column == 3) // Modified column
        {
            var fileX = itemX.Tag as FileResultItem;
            var fileY = itemY.Tag as FileResultItem;
            if (fileX != null && fileY != null)
            {
                result = fileX.Modified.CompareTo(fileY.Modified);
            }
        }
        else if (Column == 4) // Created column
        {
            var fileX = itemX.Tag as FileResultItem;
            var fileY = itemY.Tag as FileResultItem;
            if (fileX != null && fileY != null)
            {
                result = fileX.Created.CompareTo(fileY.Created);
            }
        }
        else // String columns (Name, Path)
        {
            result = string.Compare(itemX.SubItems[Column].Text, itemY.SubItems[Column].Text);
        }

        // Reverse if descending
        if (Order == SortOrder.Descending)
            result = -result;

        return result;
    }
}
