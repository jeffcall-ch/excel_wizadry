// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents a file or directory entry in the file system.
/// </summary>
public sealed class FileSystemEntry
{
    /// <summary>Full path to the file or directory.</summary>
    public required string FullPath { get; set; }

    /// <summary>Display name shown in the file list.</summary>
    public required string Name { get; set; }

    /// <summary>Whether this entry represents a directory.</summary>
    public required bool IsDirectory { get; init; }

    /// <summary>File size in bytes. Zero for directories.</summary>
    public long Size { get; init; }

    /// <summary>Last modification timestamp.</summary>
    public DateTimeOffset DateModified { get; init; }

    /// <summary>Shell-provided type description (e.g., "Text Document").</summary>
    public string TypeDescription { get; set; } = string.Empty;

    /// <summary>File extension including the leading dot.</summary>
    public string Extension { get; set; } = string.Empty;

    /// <summary>File attributes from the file system.</summary>
    public FileAttributes Attributes { get; init; }

    /// <summary>Index of the shell icon for this entry (from SHGetFileInfo).</summary>
    public int IconIndex { get; set; }

    /// <summary>OneDrive cloud status for this file.</summary>
    public CloudFileStatus CloudStatus { get; set; } = CloudFileStatus.NotApplicable;

    /// <summary>Whether the entry is hidden.</summary>
    public bool IsHidden => (Attributes & FileAttributes.Hidden) != 0;

    /// <summary>Whether the entry is a system file.</summary>
    public bool IsSystem => (Attributes & FileAttributes.System) != 0;

    /// <summary>Parent directory path.</summary>
    public string? ParentPath => Path.GetDirectoryName(FullPath);
}
