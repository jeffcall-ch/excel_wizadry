// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Text;
using FileExplorer.Core.Interfaces;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Preview pane content generation service for image, text, PDF, and shell thumbnails.
/// </summary>
public sealed class PreviewService : IPreviewService
{
    private readonly ILogger<PreviewService> _logger;
    private const int MaxTextPreviewBytes = 64 * 1024; // 64 KB

    private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".ico", ".tiff", ".tif"
    };

    private static readonly HashSet<string> TextExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".txt", ".md", ".csv", ".log", ".xml", ".json", ".yaml", ".yml",
        ".ini", ".cfg", ".conf", ".bat", ".cmd", ".ps1", ".sh",
        ".cs", ".py", ".js", ".ts", ".html", ".htm", ".css", ".sql",
        ".cpp", ".h", ".c", ".java", ".rs", ".go", ".rb", ".php",
        ".toml", ".env", ".gitignore", ".editorconfig", ".sln", ".csproj", ".props"
    };

    /// <summary>Initializes a new instance of the <see cref="PreviewService"/> class.</summary>
    public PreviewService(ILogger<PreviewService> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public async Task<PreviewContent> GetPreviewAsync(string filePath, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return new PreviewContent { Type = PreviewType.None };
            }

            var ext = Path.GetExtension(filePath);

            if (ImageExtensions.Contains(ext))
            {
                return new PreviewContent { Type = PreviewType.Image, FilePath = filePath };
            }

            if (string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                return new PreviewContent { Type = PreviewType.Pdf, FilePath = filePath };
            }

            if (TextExtensions.Contains(ext))
            {
                var text = await ReadTextPreviewAsync(filePath, cancellationToken);
                return new PreviewContent { Type = PreviewType.Text, TextContent = text, FilePath = filePath };
            }

            // Fallback to shell thumbnail
            return new PreviewContent { Type = PreviewType.ShellThumbnail, FilePath = filePath };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to generate preview for {Path}", filePath);
            return new PreviewContent { Type = PreviewType.Unsupported };
        }
    }

    private static async Task<string> ReadTextPreviewAsync(string filePath, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.Asynchronous);
        var buffer = new byte[Math.Min(MaxTextPreviewBytes, stream.Length)];
        int bytesRead = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken);

        // Detect encoding: BOM first, then fall back to UTF8
        var encoding = DetectEncoding(buffer, bytesRead);
        return encoding.GetString(buffer, 0, bytesRead);
    }

    private static Encoding DetectEncoding(byte[] buffer, int length)
    {
        if (length >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
            return Encoding.UTF8;
        if (length >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE)
            return Encoding.Unicode;
        if (length >= 2 && buffer[0] == 0xFE && buffer[1] == 0xFF)
            return Encoding.BigEndianUnicode;

        // Heuristic: check for null bytes (likely binary or UTF-16 without BOM)
        for (int i = 0; i < Math.Min(length, 1024); i++)
        {
            if (buffer[i] == 0) return Encoding.Unicode;
        }

        return Encoding.UTF8;
    }
}
