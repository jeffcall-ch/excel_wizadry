// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Service for preview pane content generation.
/// </summary>
public interface IPreviewService
{
    /// <summary>Gets preview content for the specified file path.</summary>
    Task<PreviewContent> GetPreviewAsync(string filePath, CancellationToken cancellationToken = default);
}

/// <summary>
/// Represents preview content for the preview pane.
/// </summary>
public sealed class PreviewContent
{
    /// <summary>Type of preview content.</summary>
    public required PreviewType Type { get; init; }

    /// <summary>Image source for image previews.</summary>
    public object? ImageSource { get; init; }

    /// <summary>Text content for text file previews.</summary>
    public string? TextContent { get; init; }

    /// <summary>File path for shell thumbnail fallback.</summary>
    public string? FilePath { get; init; }
}

/// <summary>
/// Types of preview content.
/// </summary>
public enum PreviewType
{
    None,
    Image,
    Text,
    Pdf,
    ShellThumbnail,
    Unsupported
}
