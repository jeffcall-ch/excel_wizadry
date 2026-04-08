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
    public required bool IsDirectory { get; set; }

    /// <summary>File size in bytes. Zero for directories.</summary>
    public long Size { get; set; }

    /// <summary>Last modification timestamp.</summary>
    public DateTimeOffset DateModified { get; set; }

    /// <summary>Shell-provided type description (e.g., "Text Document").</summary>
    public string TypeDescription { get; set; } = string.Empty;

    /// <summary>File extension including the leading dot.</summary>
    public string Extension { get; set; } = string.Empty;

    /// <summary>File attributes from the file system.</summary>
    public FileAttributes Attributes { get; set; }

    /// <summary>NTFS reparse tag, when <see cref="Attributes"/> includes a reparse point.</summary>
    public uint ReparseTag { get; set; }

    /// <summary>Index of the shell icon for this entry (from SHGetFileInfo).</summary>
    public int IconIndex { get; set; }

    /// <summary>OneDrive cloud status for this file.</summary>
    public CloudFileStatus CloudStatus { get; set; } = CloudFileStatus.NotApplicable;

    /// <summary>Extracted project title from file content anchors.</summary>
    public string ExtractedProject { get; set; } = string.Empty;

    /// <summary>Extracted revision from file content anchors.</summary>
    public string ExtractedRevision { get; set; } = string.Empty;

    /// <summary>Extracted document name from file content anchors.</summary>
    public string ExtractedDocumentName { get; set; } = string.Empty;

    /// <summary>Whether the entry is hidden.</summary>
    public bool IsHidden => (Attributes & FileAttributes.Hidden) != 0;

    /// <summary>Whether the entry is a system file.</summary>
    public bool IsSystem => (Attributes & FileAttributes.System) != 0;

    /// <summary>Parent directory path.</summary>
    public string? ParentPath => Path.GetDirectoryName(FullPath);

    /// <summary>Stable file identity key used for diff/reconciliation.</summary>
    public string FileKey => $"{NormalizePathForKey(FullPath)}|{(IsDirectory ? 1 : 0)}|{Size}|{DateModified.UtcTicks}";

    /// <inheritdoc/>
    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(this, obj))
            return true;

        if (obj is not FileSystemEntry other)
            return false;

        return IsDirectory == other.IsDirectory &&
               Size == other.Size &&
               DateModified.UtcTicks == other.DateModified.UtcTicks &&
               string.Equals(NormalizePathForKey(FullPath), NormalizePathForKey(other.FullPath), StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(NormalizePathForKey(FullPath), StringComparer.OrdinalIgnoreCase);
        hash.Add(IsDirectory);
        hash.Add(Size);
        hash.Add(DateModified.UtcTicks);
        return hash.ToHashCode();
    }

    private static string NormalizePathForKey(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        var candidate = path.Trim();
        if (candidate.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            candidate = @"\\" + candidate[8..];
        else if (candidate.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            candidate = candidate[4..];

        try
        {
            candidate = Path.GetFullPath(candidate);
        }
        catch
        {
            // Keep best-effort key when normalization fails.
        }

        return candidate.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}
