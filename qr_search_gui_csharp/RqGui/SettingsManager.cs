using System;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace RqGui
{
    /// <summary>
    /// Application settings data model.
    /// </summary>
    public class AppSettings
    {
        public string RqExePath { get; set; } = "";
        public string LastSearchPath { get; set; } = "";
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
    }
}
