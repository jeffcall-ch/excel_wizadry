// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Provides file search capabilities using the rq engine.
/// </summary>
public interface ISearchService
{
    /// <summary>
    /// Searches for files and folders matching the given term under the specified directory.
    /// </summary>
    /// <param name="directory">Root directory to search in.</param>
    /// <param name="searchTerm">Search term (filename pattern).</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>List of full paths that matched.</returns>
    Task<IReadOnlyList<string>> SearchAsync(string directory, string searchTerm, CancellationToken cancellationToken = default);
}
