using System;
using System.IO;

namespace RqGui
{
    /// <summary>
    /// Simple file-based logger for diagnostics and troubleshooting.
    /// Thread-safe and designed to never crash the application.
    /// </summary>
    public static class Logger
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RqGui",
            "rqgui.log"
        );
        
        private static readonly object lockObject = new object();
        
        /// <summary>
        /// Log an informational message.
        /// </summary>
        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }
        
        /// <summary>
        /// Log an error message.
        /// </summary>
        public static void LogError(string message)
        {
            Log("ERROR", message);
        }
        
        /// <summary>
        /// Log a warning message.
        /// </summary>
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
        
        /// <summary>
        /// Log an exception with context.
        /// </summary>
        public static void LogException(string context, Exception ex)
        {
            LogError($"{context}: {ex.GetType().Name} - {ex.Message}\n{ex.StackTrace}");
        }
    }
}
