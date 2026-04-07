// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Persists the current folder path and display title for each open tab in SQLite.
/// </summary>
public interface ITabPathStateService
{
    /// <summary>Upserts the current path/title state for a tab.</summary>
    Task UpsertCurrentPathAsync(Guid tabId, string currentPath, string displayTitle, CancellationToken cancellationToken = default);

    /// <summary>Gets the persisted path/title state for a tab.</summary>
    Task<TabPathStateRecord?> GetByTabIdAsync(Guid tabId, CancellationToken cancellationToken = default);

    /// <summary>Deletes persisted state for a closed tab.</summary>
    Task DeleteByTabIdAsync(Guid tabId, CancellationToken cancellationToken = default);
}
