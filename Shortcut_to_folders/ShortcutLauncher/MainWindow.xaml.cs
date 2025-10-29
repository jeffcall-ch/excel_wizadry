using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Extensions.Configuration;

namespace ShortcutLauncher
{
    public partial class MainWindow : Window
    {
        private readonly string _shortcutsFolder;
        private readonly int _maxItems;
        private bool _menuIsOpen = false;

        public MainWindow()
        {
            InitializeComponent();
            
            // Load configuration
            var config = LoadConfiguration();
            _shortcutsFolder = Environment.ExpandEnvironmentVariables(
                config["ShortcutsFolder"] ?? 
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Shortcuts_to_folders")
            );
            _maxItems = int.TryParse(config["MaxItems"], out int max) ? max : 50;

            // Ensure the folder exists
            if (!Directory.Exists(_shortcutsFolder))
            {
                try
                {
                    Directory.CreateDirectory(_shortcutsFolder);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to create shortcuts folder: {ex.Message}");
                }
            }
        }

        private Dictionary<string, string> LoadConfiguration()
        {
            var config = new Dictionary<string, string>();
            var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

            if (System.IO.File.Exists(configPath))
            {
                try
                {
                    var configuration = new ConfigurationBuilder()
                        .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                        .AddJsonFile("appsettings.json", optional: true)
                        .Build();

                    config["ShortcutsFolder"] = configuration["ShortcutsFolder"] ?? "";
                    config["MaxItems"] = configuration["MaxItems"] ?? "50";
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to load config: {ex.Message}");
                }
            }

            return config;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Window starts normally visible
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // When user closes the window (X button or Alt+F4), ask what to do
            var result = MessageBox.Show(
                "Do you want to minimize to taskbar or exit completely?\n\n" +
                "• Click 'Yes' to minimize (app stays in taskbar)\n" +
                "• Click 'No' to exit completely\n" +
                "• Click 'Cancel' to keep window open",
                "Close Shortcut Launcher",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                // Minimize to taskbar
                e.Cancel = true; // Cancel the close
                this.WindowState = WindowState.Minimized;
            }
            else if (result == MessageBoxResult.No)
            {
                // Exit completely - allow close to proceed
                Application.Current.Shutdown();
            }
            else
            {
                // Cancel - keep window open
                e.Cancel = true;
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            // When restored from minimized (clicked on taskbar), show menu
            if (this.WindowState == WindowState.Normal && !_menuIsOpen)
            {
                BuildMenu();
                ShowMenuAtCursor();
            }
        }

        private void ShowMenuButton_Click(object sender, RoutedEventArgs e)
        {
            BuildMenu();
            ShowMenuAtCursor();
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Directory.Exists(_shortcutsFolder))
                {
                    Directory.CreateDirectory(_shortcutsFolder);
                }

                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{_shortcutsFolder}\"")
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open folder:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl+R to refresh
            if (e.Key == Key.R && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (_menuIsOpen)
                {
                    BuildMenu();
                    e.Handled = true;
                }
            }
        }

        private void BuildMenu()
        {
            ShortcutMenu.Items.Clear();

            if (!Directory.Exists(_shortcutsFolder))
            {
                var notFoundItem = new MenuItem
                {
                    Header = "Shortcuts folder not found",
                    IsEnabled = false
                };
                ShortcutMenu.Items.Add(notFoundItem);
                ShortcutMenu.Items.Add(new Separator());
            }
            else
            {
                try
                {
                    var shortcuts = Directory.GetFiles(_shortcutsFolder, "*.lnk")
                        .Select(f => new FileInfo(f))
                        .OrderBy(f => f.Name)
                        .Take(_maxItems)
                        .ToList();

                    if (shortcuts.Count == 0)
                    {
                        var emptyItem = new MenuItem
                        {
                            Header = "No shortcuts found",
                            IsEnabled = false
                        };
                        ShortcutMenu.Items.Add(emptyItem);
                        ShortcutMenu.Items.Add(new Separator());
                    }
                    else
                    {
                        int itemCount = 0;
                        foreach (var shortcut in shortcuts)
                        {
                            var menuItem = new MenuItem
                            {
                                Header = Path.GetFileNameWithoutExtension(shortcut.Name),
                                Tag = shortcut.FullName
                            };

                            // Try to get the target path for tooltip
                            try
                            {
                                string targetPath = GetShortcutTarget(shortcut.FullName);
                                if (!string.IsNullOrEmpty(targetPath))
                                {
                                    menuItem.ToolTip = targetPath;
                                }
                            }
                            catch { }

                            menuItem.Click += ShortcutItem_Click;
                            ShortcutMenu.Items.Add(menuItem);

                            itemCount++;
                            // Add separator after every 10 items for readability
                            if (itemCount % 10 == 0 && itemCount < shortcuts.Count)
                            {
                                ShortcutMenu.Items.Add(new Separator());
                            }
                        }

                        ShortcutMenu.Items.Add(new Separator());
                    }
                }
                catch (Exception ex)
                {
                    var errorItem = new MenuItem
                    {
                        Header = $"Error reading shortcuts: {ex.Message}",
                        IsEnabled = false
                    };
                    ShortcutMenu.Items.Add(errorItem);
                    ShortcutMenu.Items.Add(new Separator());
                }
            }

            // Add bottom menu items
            var refreshItem = new MenuItem { Header = "🔄 Refresh" };
            refreshItem.Click += RefreshItem_Click;
            ShortcutMenu.Items.Add(refreshItem);

            var openFolderItem = new MenuItem { Header = "📁 Open Shortcuts Folder" };
            openFolderItem.Click += OpenFolderItem_Click;
            ShortcutMenu.Items.Add(openFolderItem);

            ShortcutMenu.Items.Add(new Separator());

            var exitItem = new MenuItem { Header = "❌ Exit" };
            exitItem.Click += ExitItem_Click;
            ShortcutMenu.Items.Add(exitItem);
        }

        private string GetShortcutTarget(string shortcutPath)
        {
            try
            {
                // Use dynamic COM invocation for .NET 8 compatibility
                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType == null) return string.Empty;
                
                dynamic shell = Activator.CreateInstance(shellType);
                dynamic shortcut = shell.CreateShortcut(shortcutPath);
                string targetPath = shortcut.TargetPath;
                
                // Release COM objects
                if (shortcut != null) Marshal.ReleaseComObject(shortcut);
                if (shell != null) Marshal.ReleaseComObject(shell);
                
                return targetPath ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void ShowMenuAtCursor()
        {
            _menuIsOpen = true;

            // Get cursor position
            var point = System.Windows.Forms.Cursor.Position;
            
            // Convert to WPF coordinates
            var dpi = VisualTreeHelper.GetDpi(this);
            var x = point.X / dpi.DpiScaleX;
            var y = point.Y / dpi.DpiScaleY;

            // Position the window near the cursor
            this.Left = x;
            this.Top = y;

            // Open the context menu
            ShortcutMenu.PlacementTarget = this;
            ShortcutMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse;
            ShortcutMenu.IsOpen = true;
        }

        private void ShortcutMenu_Closed(object sender, RoutedEventArgs e)
        {
            _menuIsOpen = false;
            
            // Just minimize to taskbar, don't hide
            if (this.WindowState != WindowState.Minimized)
            {
                this.WindowState = WindowState.Minimized;
            }
        }

        private void ShortcutItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is string linkPath)
            {
                try
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", $"\"{linkPath}\"")
                    {
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to open shortcut:\n{ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }

                ShortcutMenu.IsOpen = false;
            }
        }

        private void RefreshItem_Click(object sender, RoutedEventArgs e)
        {
            BuildMenu();
        }

        private void OpenFolderItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Directory.Exists(_shortcutsFolder))
                {
                    Directory.CreateDirectory(_shortcutsFolder);
                }

                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{_shortcutsFolder}\"")
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open folder:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            ShortcutMenu.IsOpen = false;
        }

        private void ExitItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
