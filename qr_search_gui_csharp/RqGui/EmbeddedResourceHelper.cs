using System;
using System.IO;
using System.Reflection;

namespace RqGui
{
    /// <summary>
    /// Helper class to extract embedded resources (like rq.exe) from the assembly.
    /// </summary>
    public static class EmbeddedResourceHelper
    {
        /// <summary>
        /// Extracts the embedded rq.exe to a temporary location and returns the path.
        /// </summary>
        public static string ExtractRqExe()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "RqGui.Resources.rq.exe";

                // Create temp directory for extracted exe
                var tempDir = Path.Combine(Path.GetTempPath(), "RqGui");
                Directory.CreateDirectory(tempDir);

                var targetPath = Path.Combine(tempDir, "rq.exe");

                // Check if already extracted and valid
                if (File.Exists(targetPath))
                {
                    // Verify it's not corrupted by checking file size
                    using (var resourceStream = assembly.GetManifestResourceStream(resourceName))
                    {
                        if (resourceStream != null)
                        {
                            var fileInfo = new FileInfo(targetPath);
                            if (fileInfo.Length == resourceStream.Length)
                            {
                                Logger.LogInfo($"Using existing rq.exe from: {targetPath}");
                                return targetPath;
                            }
                        }
                    }
                }

                // Extract the embedded resource
                using (var resourceStream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (resourceStream == null)
                    {
                        throw new FileNotFoundException($"Embedded resource '{resourceName}' not found in assembly.");
                    }

                    using (var fileStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write))
                    {
                        resourceStream.CopyTo(fileStream);
                    }
                }

                Logger.LogInfo($"Extracted rq.exe to: {targetPath}");
                return targetPath;
            }
            catch (Exception ex)
            {
                Logger.LogException("Failed to extract embedded rq.exe", ex);
                throw;
            }
        }

        /// <summary>
        /// Gets the path to rq.exe - either from settings or from embedded resource.
        /// </summary>
        public static string GetRqExePath()
        {
            var settings = SettingsManager.LoadSettings();

            // If user has configured a custom path and it exists, use it
            if (!string.IsNullOrWhiteSpace(settings.RqExePath) && File.Exists(settings.RqExePath))
            {
                Logger.LogInfo($"Using configured rq.exe path: {settings.RqExePath}");
                return settings.RqExePath;
            }

            // Otherwise, extract and use the embedded version
            Logger.LogInfo("No valid rq.exe path configured, using embedded version");
            var embeddedPath = ExtractRqExe();

            // Update settings to point to embedded version
            settings.RqExePath = embeddedPath;
            SettingsManager.SaveSettings(settings);

            return embeddedPath;
        }
    }
}
