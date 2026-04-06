// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Background deep-indexing coordinator for favorites/fly-out roots.
/// </summary>
public interface IDeepIndexingService
{
    /// <summary>Human-readable warning when the indexing engine is unavailable.</summary>
    string? EngineWarning { get; }

    /// <summary>True while a deep-indexing run is active.</summary>
    bool IsRunning { get; }

    /// <summary>
    /// Starts or refreshes background recursive indexing for the provided root folders.
    /// </summary>
    Task StartOrRefreshAsync(
        IReadOnlyList<string> favoriteRootPaths,
        CancellationToken cancellationToken = default);
}
