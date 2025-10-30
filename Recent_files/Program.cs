using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

namespace RecentLauncher;

// ============================================================================
// ENTRY POINT
// ============================================================================
static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}

// ============================================================================
// DATA MODEL
// ============================================================================
/// <summary>
/// Represents a single bookmark (file or folder).
/// </summary>
class Bookmark
{
    public string DisplayName { get; set; } = "";
    public string FullPath { get; set; } = "";
    public string Type { get; set; } = ""; // "File" or "Folder"

    public bool Exists()
    {
        if (Type == "File")
            return File.Exists(FullPath);
        else if (Type == "Folder")
            return Directory.Exists(FullPath);
        return false;
    }
}

/// <summary>
/// Persistent data structure including bookmarks and window bounds.
/// </summary>
class AppData
{
    public List<Bookmark> Bookmarks { get; set; } = new();
    public Rectangle? WindowBounds { get; set; }
}

// ============================================================================
// MAIN FORM
// ============================================================================
class MainForm : Form
{
    // UI Controls
    private Label titleLabel = null!;
    private ListBox bookmarkListBox = null!;
    private Button addButton = null!;
    private Button openButton = null!;
    private Button openLocationButton = null!;
    private Button removeButton = null!;
    private Label countLabel = null!;
    private ContextMenuStrip contextMenu = null!;

    // Data
    private List<Bookmark> bookmarks = new();
    private string dataFilePath = null!;

    public MainForm()
    {
        InitializeComponents();
        SetupEventHandlers();
        LoadData();
        RefreshList();
    }

    // ========================================================================
    // UI INITIALIZATION
    // ========================================================================
    private void InitializeComponents()
    {
        // Form properties
        Text = "Recent Launcher";
        Size = new Size(700, 450);
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(400, 300);

        // Title label
        titleLabel = new Label
        {
            Text = "Recent Launcher",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(10, 10)
        };

        // ListBox for bookmarks
        bookmarkListBox = new ListBox
        {
            Location = new Point(10, 50),
            Size = new Size(660, 300),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            SelectionMode = SelectionMode.One,
            AllowDrop = true // Enable drag-and-drop
        };

        // Bottom panel
        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            Padding = new Padding(10, 5, 10, 5)
        };

        // Buttons
        addButton = new Button
        {
            Text = "Add...",
            Location = new Point(10, 10),
            Width = 100,
            Height = 30
        };

        openButton = new Button
        {
            Text = "Open",
            Location = new Point(120, 10),
            Width = 100,
            Height = 30
        };

        openLocationButton = new Button
        {
            Text = "Open Location",
            Location = new Point(230, 10),
            Width = 120,
            Height = 30
        };

        removeButton = new Button
        {
            Text = "Remove",
            Location = new Point(360, 10),
            Width = 100,
            Height = 30
        };

        countLabel = new Label
        {
            Text = "0 items",
            AutoSize = true,
            Location = new Point(600, 15),
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };

        // Add controls to bottom panel
        bottomPanel.Controls.Add(addButton);
        bottomPanel.Controls.Add(openButton);
        bottomPanel.Controls.Add(openLocationButton);
        bottomPanel.Controls.Add(removeButton);
        bottomPanel.Controls.Add(countLabel);

        // Context menu
        contextMenu = new ContextMenuStrip();
        contextMenu.Items.Add("Open", null, (s, e) => OpenSelected());
        contextMenu.Items.Add("Open Location", null, (s, e) => OpenLocation());
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add("Rename...", null, (s, e) => RenameSelected());
        contextMenu.Items.Add("Remove", null, (s, e) => RemoveSelected());

        bookmarkListBox.ContextMenuStrip = contextMenu;

        // Add controls to form
        Controls.Add(titleLabel);
        Controls.Add(bookmarkListBox);
        Controls.Add(bottomPanel);

        // Data file path
        var appDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RecentLauncher"
        );
        Directory.CreateDirectory(appDataFolder);
        dataFilePath = Path.Combine(appDataFolder, "bookmarks.json");
    }

    // ========================================================================
    // EVENT HANDLERS SETUP
    // ========================================================================
    private void SetupEventHandlers()
    {
        // Button events
        addButton.Click += (s, e) => ShowAddDialog();
        openButton.Click += (s, e) => OpenSelected();
        openLocationButton.Click += (s, e) => OpenLocation();
        removeButton.Click += (s, e) => RemoveSelected();

        // ListBox events
        bookmarkListBox.DoubleClick += (s, e) => OpenSelected();
        bookmarkListBox.SelectedIndexChanged += (s, e) => UpdateButtonStates();
        bookmarkListBox.KeyDown += ListBox_KeyDown;
        
        // Drag-and-drop events
        bookmarkListBox.DragEnter += ListBox_DragEnter;
        bookmarkListBox.DragDrop += ListBox_DragDrop;

        // Form events
        FormClosing += (s, e) => SaveData();
        KeyPreview = true;
        KeyDown += Form_KeyDown;
    }

    // ========================================================================
    // KEYBOARD SHORTCUTS
    // ========================================================================
    private void Form_KeyDown(object? sender, KeyEventArgs e)
    {
        // Handle form-level keyboard shortcuts
        if (e.Control && e.KeyCode == Keys.N)
        {
            ShowAddDialog();
            e.Handled = true;
        }
        else if (e.Control && e.KeyCode == Keys.O)
        {
            OpenSelected();
            e.Handled = true;
        }
        else if (e.Control && e.KeyCode == Keys.L)
        {
            OpenLocation();
            e.Handled = true;
        }
        else if (e.Control && e.KeyCode == Keys.S)
        {
            SaveData();
            e.Handled = true;
        }
    }

    private void ListBox_KeyDown(object? sender, KeyEventArgs e)
    {
        // Handle ListBox-specific keyboard shortcuts
        if (e.KeyCode == Keys.Enter)
        {
            OpenSelected();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Delete)
        {
            RemoveSelected();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.F2)
        {
            RenameSelected();
            e.Handled = true;
        }
    }

    // ========================================================================
    // DRAG-AND-DROP
    // ========================================================================
    private void ListBox_DragEnter(object? sender, DragEventArgs e)
    {
        // Accept file/folder drops
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            e.Effect = DragDropEffects.Copy;
        }
        else
        {
            e.Effect = DragDropEffects.None;
        }
    }

    private void ListBox_DragDrop(object? sender, DragEventArgs e)
    {
        // Handle dropped files/folders
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] paths)
        {
            foreach (var path in paths)
            {
                AddPath(path);
            }
        }
    }

    // ========================================================================
    // CORE FUNCTIONALITY
    // ========================================================================

    /// <summary>
    /// Shows dialog to add a new bookmark (file or folder).
    /// </summary>
    private void ShowAddDialog()
    {
        // Simple choice dialog
        var choice = MessageBox.Show(
            "Choose:\nYes = Add Folder\nNo = Add File",
            "Add Bookmark",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question
        );

        if (choice == DialogResult.Yes)
        {
            // Add folder
            using var dialog = new FolderBrowserDialog();
            dialog.Description = "Select a folder to add";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                AddPath(dialog.SelectedPath);
            }
        }
        else if (choice == DialogResult.No)
        {
            // Add file
            using var dialog = new OpenFileDialog();
            dialog.Title = "Select a file to add";
            dialog.Filter = "All Files (*.*)|*.*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                AddPath(dialog.FileName);
            }
        }
    }

    /// <summary>
    /// Adds a path (file or folder) to bookmarks.
    /// </summary>
    private void AddPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        // Check for duplicates
        if (bookmarks.Any(b => b.FullPath.Equals(path, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show(
                "This path is already in the list.",
                "Duplicate",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
            return;
        }

        // Determine type and create bookmark
        string type;
        string displayName;

        if (Directory.Exists(path))
        {
            type = "Folder";
            displayName = new DirectoryInfo(path).Name;
        }
        else if (File.Exists(path))
        {
            type = "File";
            displayName = Path.GetFileName(path);
        }
        else
        {
            MessageBox.Show(
                "The selected path does not exist.",
                "Invalid Path",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }

        var bookmark = new Bookmark
        {
            DisplayName = displayName,
            FullPath = path,
            Type = type
        };

        bookmarks.Add(bookmark);
        SaveData();
        RefreshList();
    }

    /// <summary>
    /// Opens the currently selected bookmark.
    /// </summary>
    private void OpenSelected()
    {
        if (bookmarkListBox.SelectedIndex < 0)
            return;

        var bookmark = bookmarks[bookmarkListBox.SelectedIndex];

        if (!bookmark.Exists())
        {
            MessageBox.Show(
                $"The path no longer exists:\n{bookmark.FullPath}",
                "Path Not Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = bookmark.FullPath,
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error opening path:\n{ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    /// <summary>
    /// Opens the parent folder and selects the item.
    /// </summary>
    private void OpenLocation()
    {
        if (bookmarkListBox.SelectedIndex < 0)
            return;

        var bookmark = bookmarks[bookmarkListBox.SelectedIndex];

        try
        {
            if (bookmark.Type == "File" && File.Exists(bookmark.FullPath))
            {
                // Open folder and select the file
                Process.Start("explorer.exe", $"/select,\"{bookmark.FullPath}\"");
            }
            else if (bookmark.Type == "Folder" && Directory.Exists(bookmark.FullPath))
            {
                // Open the folder itself
                var parentPath = Path.GetDirectoryName(bookmark.FullPath);
                if (!string.IsNullOrEmpty(parentPath) && Directory.Exists(parentPath))
                {
                    Process.Start("explorer.exe", $"/select,\"{bookmark.FullPath}\"");
                }
                else
                {
                    // If no parent, just open the folder
                    Process.Start("explorer.exe", bookmark.FullPath);
                }
            }
            else
            {
                MessageBox.Show(
                    $"The path no longer exists:\n{bookmark.FullPath}",
                    "Path Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error opening location:\n{ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    /// <summary>
    /// Removes the currently selected bookmark.
    /// </summary>
    private void RemoveSelected()
    {
        if (bookmarkListBox.SelectedIndex < 0)
            return;

        var index = bookmarkListBox.SelectedIndex;
        bookmarks.RemoveAt(index);
        SaveData();
        RefreshList();

        // Try to keep selection nearby
        if (bookmarks.Count > 0)
        {
            bookmarkListBox.SelectedIndex = Math.Min(index, bookmarks.Count - 1);
        }
    }

    /// <summary>
    /// Renames the display name of the selected bookmark.
    /// </summary>
    private void RenameSelected()
    {
        if (bookmarkListBox.SelectedIndex < 0)
            return;

        var bookmark = bookmarks[bookmarkListBox.SelectedIndex];
        var newName = Prompt.ShowDialog("Enter new display name:", "Rename", bookmark.DisplayName);

        if (!string.IsNullOrWhiteSpace(newName))
        {
            bookmark.DisplayName = newName;
            SaveData();
            RefreshList();
        }
    }

    /// <summary>
    /// Refreshes the ListBox display with current bookmarks.
    /// </summary>
    private void RefreshList()
    {
        var selectedIndex = bookmarkListBox.SelectedIndex;
        bookmarkListBox.Items.Clear();

        foreach (var bookmark in bookmarks)
        {
            var display = bookmark.DisplayName;
            if (!bookmark.Exists())
            {
                display += " (missing)";
            }
            bookmarkListBox.Items.Add(display);
        }

        countLabel.Text = $"{bookmarks.Count} item{(bookmarks.Count != 1 ? "s" : "")}";

        // Restore selection
        if (selectedIndex >= 0 && selectedIndex < bookmarks.Count)
        {
            bookmarkListBox.SelectedIndex = selectedIndex;
        }

        UpdateButtonStates();
    }

    /// <summary>
    /// Updates the enabled state of buttons based on selection and item validity.
    /// </summary>
    private void UpdateButtonStates()
    {
        var hasSelection = bookmarkListBox.SelectedIndex >= 0;
        var exists = hasSelection && bookmarks[bookmarkListBox.SelectedIndex].Exists();

        openButton.Enabled = hasSelection && exists;
        openLocationButton.Enabled = hasSelection;
        removeButton.Enabled = hasSelection;
    }

    // ========================================================================
    // PERSISTENCE
    // ========================================================================

    /// <summary>
    /// Loads bookmarks and window bounds from JSON file.
    /// </summary>
    private void LoadData()
    {
        if (!File.Exists(dataFilePath))
            return;

        try
        {
            var json = File.ReadAllText(dataFilePath);
            var data = JsonSerializer.Deserialize<AppData>(json);

            if (data != null)
            {
                bookmarks = data.Bookmarks ?? new();

                // Restore window bounds
                if (data.WindowBounds.HasValue)
                {
                    var bounds = data.WindowBounds.Value;
                    // Ensure the window is visible on screen
                    if (IsVisibleOnAnyScreen(bounds))
                    {
                        StartPosition = FormStartPosition.Manual;
                        Bounds = bounds;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error loading data:\n{ex.Message}",
                "Load Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    /// <summary>
    /// Saves bookmarks and window bounds to JSON file.
    /// </summary>
    private void SaveData()
    {
        try
        {
            var data = new AppData
            {
                Bookmarks = bookmarks,
                WindowBounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(data, options);
            File.WriteAllText(dataFilePath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error saving data:\n{ex.Message}",
                "Save Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }
    }

    /// <summary>
    /// Checks if a rectangle is visible on any screen.
    /// </summary>
    private bool IsVisibleOnAnyScreen(Rectangle rect)
    {
        foreach (var screen in Screen.AllScreens)
        {
            if (screen.WorkingArea.IntersectsWith(rect))
                return true;
        }
        return false;
    }
}

// ============================================================================
// HELPER: SIMPLE INPUT DIALOG
// ============================================================================
static class Prompt
{
    public static string ShowDialog(string text, string caption, string defaultValue = "")
    {
        Form prompt = new Form()
        {
            Width = 400,
            Height = 150,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Text = caption,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false,
            MinimizeBox = false
        };

        Label textLabel = new Label() { Left = 20, Top = 20, Text = text, Width = 340 };
        TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 340, Text = defaultValue };
        Button confirmation = new Button() { Text = "OK", Left = 200, Width = 80, Top = 80, DialogResult = DialogResult.OK };
        Button cancel = new Button() { Text = "Cancel", Left = 290, Width = 80, Top = 80, DialogResult = DialogResult.Cancel };

        confirmation.Click += (sender, e) => { prompt.Close(); };
        cancel.Click += (sender, e) => { prompt.Close(); };

        prompt.Controls.Add(textLabel);
        prompt.Controls.Add(textBox);
        prompt.Controls.Add(confirmation);
        prompt.Controls.Add(cancel);
        prompt.AcceptButton = confirmation;
        prompt.CancelButton = cancel;

        return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : defaultValue;
    }
}
