// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// SQLite store for current folder path per tab.
/// </summary>
public sealed class SqliteTabPathStateService : ITabPathStateService
{
    private readonly IIdentityCacheService _identityCacheService;
    private readonly ILogger<SqliteTabPathStateService> _logger;
    private readonly SemaphoreSlim _initGate = new(1, 1);
    private volatile bool _initialized;

    public SqliteTabPathStateService(
        IIdentityCacheService identityCacheService,
        ILogger<SqliteTabPathStateService> logger)
    {
        _identityCacheService = identityCacheService;
        _logger = logger;
    }

    public async Task UpsertCurrentPathAsync(Guid tabId, string currentPath, string displayTitle, CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        var normalizedPath = _identityCacheService.NormalizePath(currentPath);
        if (string.IsNullOrWhiteSpace(normalizedPath))
            return;

        var safeTitle = string.IsNullOrWhiteSpace(displayTitle) ? normalizedPath : displayTitle.Trim();
        var tabKey = BuildTabKey(tabId);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            INSERT INTO TabCurrentFolderState
                (TabId, TabKey, CurrentPath, DisplayTitle, UpdatedUtcTicks)
            VALUES
                ($tabId, $tabKey, $currentPath, $displayTitle, $updatedUtcTicks)
            ON CONFLICT(TabId) DO UPDATE SET
                TabKey = excluded.TabKey,
                CurrentPath = excluded.CurrentPath,
                DisplayTitle = excluded.DisplayTitle,
                UpdatedUtcTicks = excluded.UpdatedUtcTicks;
            """;

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("$tabId", tabId.ToString("D"));
        command.Parameters.AddWithValue("$tabKey", tabKey);
        command.Parameters.AddWithValue("$currentPath", normalizedPath);
        command.Parameters.AddWithValue("$displayTitle", safeTitle);
        command.Parameters.AddWithValue("$updatedUtcTicks", DateTime.UtcNow.Ticks);

        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<TabPathStateRecord?> GetByTabIdAsync(Guid tabId, CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            SELECT TabId, TabKey, CurrentPath, DisplayTitle, UpdatedUtcTicks
            FROM TabCurrentFolderState
            WHERE TabId = $tabId;
            """;

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("$tabId", tabId.ToString("D"));

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
            return null;

        return new TabPathStateRecord(
            Guid.Parse(reader.GetString(0)),
            reader.GetString(1),
            reader.GetString(2),
            reader.GetString(3),
            reader.GetInt64(4));
    }

    public async Task DeleteByTabIdAsync(Guid tabId, CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = "DELETE FROM TabCurrentFolderState WHERE TabId = $tabId;";

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("$tabId", tabId.ToString("D"));

        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task EnsureInitializedAsync(CancellationToken cancellationToken)
    {
        if (_initialized)
            return;

        await _initGate.WaitAsync(cancellationToken);
        try
        {
            if (_initialized)
                return;

            await using var connection = CreateConnection();
            await connection.OpenAsync(cancellationToken);

            const string schemaSql = """
                CREATE TABLE IF NOT EXISTS TabCurrentFolderState (
                    TabId TEXT PRIMARY KEY,
                    TabKey TEXT NOT NULL UNIQUE,
                    CurrentPath TEXT NOT NULL,
                    DisplayTitle TEXT NOT NULL,
                    UpdatedUtcTicks INTEGER NOT NULL
                );

                CREATE INDEX IF NOT EXISTS IX_TabCurrentFolderState_UpdatedUtcTicks
                    ON TabCurrentFolderState (UpdatedUtcTicks);
                """;

            await using var command = connection.CreateCommand();
            command.CommandText = schemaSql;
            await command.ExecuteNonQueryAsync(cancellationToken);

            _initialized = true;
            _logger.LogDebug("Tab current folder state table initialized in {DbPath}", _identityCacheService.DatabasePath);
        }
        finally
        {
            _initGate.Release();
        }
    }

    private SqliteConnection CreateConnection()
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = _identityCacheService.DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared
        };

        return new SqliteConnection(builder.ToString());
    }

    private static string BuildTabKey(Guid tabId) => $"tab-{tabId:N}";
}
