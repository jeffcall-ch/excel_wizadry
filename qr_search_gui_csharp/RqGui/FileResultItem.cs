using System;
using System.IO;

namespace RqGui
{
    /// <summary>
    /// Represents a file found in search results.
    /// </summary>
    public class FileResultItem
    {
        public string FileName { get; set; } = "";
        public string FullPath { get; set; } = "";
        public string? Directory { get; set; }
        public long Size { get; set; }
        public DateTime Modified { get; set; }
        public DateTime Created { get; set; }
        public string Extension { get; set; } = "";

        /// <summary>
        /// Create a FileResultItem from a file path.
        /// Safely handles files that may no longer exist.
        /// </summary>
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
                    Created = fileInfo.CreationTime,
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
                    Created = DateTime.MinValue,
                    Extension = Path.GetExtension(path)
                };
            }
        }

        /// <summary>
        /// Get human-readable file size (e.g., "1.5 MB").
        /// </summary>
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
}
