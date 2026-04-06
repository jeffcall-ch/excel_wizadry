// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents one cached identity row in the SQLite metadata store.
/// </summary>
public sealed record FileIdentityCacheRecord
{
    /// <summary>Composite key: normalizedFullPath|fileSize|lastWriteTime.</summary>
    public required string FileKey { get; init; }

    /// <summary>Normalized parent directory path in lowercase.</summary>
    public required string ParentPath { get; init; }

    /// <summary>Extracted project title.</summary>
    public string ProjectTitle { get; init; } = string.Empty;

    /// <summary>Extracted revision value.</summary>
    public string Revision { get; init; } = string.Empty;

    /// <summary>Extracted document name.</summary>
    public string DocumentName { get; init; } = string.Empty;

    /// <summary>Source file last-write timestamp in UTC ticks.</summary>
    public long LastWriteTime { get; init; }

    /// <summary>Source file size in bytes.</summary>
    public long FileSize { get; init; }

    /// <summary>Converts this cache row to a document identity payload.</summary>
    public DocumentIdentity ToDocumentIdentity()
    {
        return new DocumentIdentity
        {
            ProjectTitle = ProjectTitle,
            Revision = Revision,
            DocumentName = DocumentName
        };
    }
}
