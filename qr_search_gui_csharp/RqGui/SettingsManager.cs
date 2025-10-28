using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

namespace RqGui
{
    /// <summary>
    /// Represents a complete search history entry with all parameters.
    /// </summary>
    public class SearchHistoryItem
    {
        public string SearchPath { get; set; } = "";
        public string SearchTerm { get; set; } = "";
        public bool UseGlob { get; set; }
        public bool UseRegex { get; set; }
        public bool CaseSensitive { get; set; }
        public bool IncludeHidden { get; set; }
        public bool ExtensionsEnabled { get; set; }
        public string Extensions { get; set; } = "";
        public bool FileTypeEnabled { get; set; }
        public string FileType { get; set; } = "";
        public bool SizeEnabled { get; set; }
        public string Size { get; set; } = "";
        public bool DateAfterEnabled { get; set; }
        public DateTime DateAfter { get; set; }
        public bool DateBeforeEnabled { get; set; }
        public DateTime DateBefore { get; set; }
        public bool MaxDepthEnabled { get; set; }
        public string MaxDepth { get; set; } = "";
        public bool ThreadsEnabled { get; set; }
        public string Threads { get; set; } = "";
        public bool TimeoutEnabled { get; set; }
        public string Timeout { get; set; } = "";
        public DateTime Timestamp { get; set; } = DateTime.Now;

        /// <summary>
        /// Get a display string for this history item.
        /// </summary>
        public string GetDisplayText()
        {
            var parts = new List<string>();
            
            if (!string.IsNullOrWhiteSpace(SearchTerm))
                parts.Add($"'{SearchTerm}'");
            
            if (!string.IsNullOrWhiteSpace(SearchPath))
            {
                var dirName = Path.GetFileName(SearchPath);
                if (string.IsNullOrWhiteSpace(dirName))
                    dirName = SearchPath;
                parts.Add($"in {dirName}");
            }
            
            if (UseGlob) parts.Add("glob");
            if (UseRegex) parts.Add("regex");
            if (CaseSensitive) parts.Add("case-sensitive");
            
            if (ExtensionsEnabled && !string.IsNullOrWhiteSpace(Extensions))
                parts.Add($"ext:{Extensions}");
            
            if (FileTypeEnabled && !string.IsNullOrWhiteSpace(FileType))
                parts.Add($"type:{FileType}");
            
            if (SizeEnabled && !string.IsNullOrWhiteSpace(Size))
                parts.Add($"size:{Size}");
            
            var display = string.Join(" | ", parts);
            if (string.IsNullOrWhiteSpace(display))
                display = "Empty search";
            
            return $"{Timestamp:yyyy-MM-dd HH:mm} - {display}";
        }

        /// <summary>
        /// Override ToString for ComboBox display.
        /// </summary>
        public override string ToString()
        {
            return GetDisplayText();
        }
    }

    /// <summary>
    /// Application settings data model.
    /// </summary>
    public class AppSettings
    {
        public string RqExePath { get; set; } = "";
        public string LastSearchPath { get; set; } = "";
        public List<SearchHistoryItem> SearchHistory { get; set; } = new List<SearchHistoryItem>();
        public int MaxHistoryItems { get; set; } = 50;
    }

    /// <summary>
    /// Manages application settings persistence.
    /// </summary>
    public static class SettingsManager
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RqGui",
            "settings.json"
        );

        /// <summary>
        /// Load settings from disk. Returns default settings if file doesn't exist or cannot be read.
        /// </summary>
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
                Logger.LogError($"Failed to load settings: {ex.Message}");
            }
            return new AppSettings();
        }

        /// <summary>
        /// Save settings to disk.
        /// </summary>
        /// <returns>True if successful, false otherwise.</returns>
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

        /// <summary>
        /// Add a search to history, maintaining the maximum history count.
        /// </summary>
        public static void AddToSearchHistory(SearchHistoryItem item)
        {
            try
            {
                var settings = LoadSettings();
                
                // Add to the beginning of the list
                settings.SearchHistory.Insert(0, item);
                
                // Trim to max items
                if (settings.SearchHistory.Count > settings.MaxHistoryItems)
                {
                    settings.SearchHistory = settings.SearchHistory.GetRange(0, settings.MaxHistoryItems);
                }
                
                SaveSettings(settings);
                Logger.LogInfo($"Added search to history: {item.GetDisplayText()}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to add search to history: {ex.Message}");
            }
        }

        /// <summary>
        /// Remove a specific search from history.
        /// </summary>
        public static void RemoveFromSearchHistory(SearchHistoryItem item)
        {
            try
            {
                var settings = LoadSettings();
                
                // Find and remove the item by comparing timestamp and key properties
                var itemToRemove = settings.SearchHistory.FirstOrDefault(h => 
                    h.Timestamp == item.Timestamp && 
                    h.SearchPath == item.SearchPath && 
                    h.SearchTerm == item.SearchTerm);
                
                if (itemToRemove != null)
                {
                    settings.SearchHistory.Remove(itemToRemove);
                    SaveSettings(settings);
                    Logger.LogInfo($"Removed search from history: {item.GetDisplayText()}");
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to remove search from history: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Clear all search history.
        /// </summary>
        public static void ClearSearchHistory()
        {
            try
            {
                var settings = LoadSettings();
                settings.SearchHistory.Clear();
                SaveSettings(settings);
                Logger.LogInfo("Search history cleared");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to clear search history: {ex.Message}");
            }
        }
    }
}
