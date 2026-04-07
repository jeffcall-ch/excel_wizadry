// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// SQLite-backed current folder state for one tab.
/// </summary>
public sealed record TabPathStateRecord(
    Guid TabId,
    string TabKey,
    string CurrentPath,
    string DisplayTitle,
    long UpdatedUtcTicks);
