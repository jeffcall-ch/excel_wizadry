// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Identity fields extracted from document content using anchor phrases.
/// </summary>
public sealed record DocumentIdentity
{
    /// <summary>An empty identity payload.</summary>
    public static DocumentIdentity Empty { get; } = new();

    /// <summary>Extracted project title.</summary>
    public string ProjectTitle { get; init; } = string.Empty;

    /// <summary>Extracted revision.</summary>
    public string Revision { get; init; } = string.Empty;

    /// <summary>Extracted document name.</summary>
    public string DocumentName { get; init; } = string.Empty;

    /// <summary>Returns true when at least one identity field is populated.</summary>
    public bool HasAnyValue =>
        !string.IsNullOrWhiteSpace(ProjectTitle) ||
        !string.IsNullOrWhiteSpace(Revision) ||
        !string.IsNullOrWhiteSpace(DocumentName);
}
