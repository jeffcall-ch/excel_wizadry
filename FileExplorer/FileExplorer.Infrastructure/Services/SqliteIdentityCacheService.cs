// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// SQLite implementation of the identity cache layer.
/// </summary>
public sealed class SqliteIdentityCacheService : IIdentityCacheService
{
    private const string DbFileName = "identity_cache.db";

    private readonly ILogger<SqliteIdentityCacheService> _logger;
    private readonly SemaphoreSlim _initGate = new(1, 1);
    private volatile bool _initialized;

    /// <summary>
    /// Initializes a new instance of the <see cref="SqliteIdentityCacheService"/> class.
    /// </summary>
    public SqliteIdentityCacheService(ILogger<SqliteIdentityCacheService> logger)
    {
        _logger = logger;
        DatabasePath = ResolveDatabasePath();
    }

    /// <inheritdoc/>
    public string DatabasePath { get; }

    /// <inheritdoc/>
    public string NormalizePath(string fullPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath))
            return string.Empty;

        var candidate = StripLongPathPrefix(fullPath.Trim().Trim('"'));

        try
        {
            candidate = Path.GetFullPath(candidate);
        }
        catch
        {
            // Keep best-effort path when path is transient or malformed.
        }

        candidate = StripLongPathPrefix(candidate)
            .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

        var root = Path.GetPathRoot(candidate);
        if (!string.IsNullOrWhiteSpace(root) &&
            !candidate.Equals(root, StringComparison.OrdinalIgnoreCase))
        {
            candidate = candidate.TrimEnd(Path.DirectorySeparatorChar);
        }

        return candidate.ToLowerInvariant();
    }

    /// <inheritdoc/>
    public string BuildFileKey(string fullPath, long fileSize, DateTimeOffset lastWriteTime)
    {
        var normalizedPath = NormalizePath(fullPath);
        return $"{normalizedPath}|{fileSize}|{lastWriteTime.ToUniversalTime().Ticks}";
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyDictionary<string, FileIdentityCacheRecord>> GetEntriesByParentPathAsync(
        string parentPath,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        var normalizedParent = NormalizePath(parentPath);
        if (string.IsNullOrWhiteSpace(normalizedParent))
            return new Dictionary<string, FileIdentityCacheRecord>(StringComparer.OrdinalIgnoreCase);

        var rows = new Dictionary<string, FileIdentityCacheRecord>(StringComparer.OrdinalIgnoreCase);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            SELECT FileKey, ProjectTitle, Revision, DocumentName, LastWriteTime, FileSize, ParentPath
            FROM FileIdentityCache
            WHERE ParentPath = $parentPath;
            """;

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("$parentPath", normalizedParent);

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var row = new FileIdentityCacheRecord
            {
                FileKey = reader.GetString(0),
                ProjectTitle = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                Revision = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                DocumentName = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                LastWriteTime = reader.GetInt64(4),
                FileSize = reader.GetInt64(5),
                ParentPath = reader.GetString(6)
            };

            rows[row.FileKey] = row;
        }

        return rows;
    }

    /// <inheritdoc/>
    public async Task<FileIdentityCacheRecord?> GetEntryByFileKeyAsync(
        string fileKey,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        if (string.IsNullOrWhiteSpace(fileKey))
            return null;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            SELECT FileKey, ProjectTitle, Revision, DocumentName, LastWriteTime, FileSize, ParentPath
            FROM FileIdentityCache
            WHERE FileKey = $fileKey;
            """;

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("$fileKey", fileKey);

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
            return null;

        return new FileIdentityCacheRecord
        {
            FileKey = reader.GetString(0),
            ProjectTitle = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
            Revision = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
            DocumentName = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
            LastWriteTime = reader.GetInt64(4),
            FileSize = reader.GetInt64(5),
            ParentPath = reader.GetString(6)
        };
    }

    /// <inheritdoc/>
    public async Task UpsertEntriesAsync(
        IReadOnlyList<FileIdentityCacheRecord> entries,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        if (entries.Count == 0)
            return;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqliteTransaction)await connection.BeginTransactionAsync(cancellationToken);

        const string sql = """
            INSERT INTO FileIdentityCache
                (FileKey, ProjectTitle, Revision, DocumentName, LastWriteTime, FileSize, ParentPath)
            VALUES
                ($fileKey, $projectTitle, $revision, $documentName, $lastWriteTime, $fileSize, $parentPath)
            ON CONFLICT(FileKey) DO UPDATE SET
                ProjectTitle = excluded.ProjectTitle,
                Revision = excluded.Revision,
                DocumentName = excluded.DocumentName,
                LastWriteTime = excluded.LastWriteTime,
                FileSize = excluded.FileSize,
                ParentPath = excluded.ParentPath;
            """;

        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;

        var fileKeyParameter = command.Parameters.Add("$fileKey", SqliteType.Text);
        var projectParameter = command.Parameters.Add("$projectTitle", SqliteType.Text);
        var revisionParameter = command.Parameters.Add("$revision", SqliteType.Text);
        var documentNameParameter = command.Parameters.Add("$documentName", SqliteType.Text);
        var lastWriteParameter = command.Parameters.Add("$lastWriteTime", SqliteType.Integer);
        var fileSizeParameter = command.Parameters.Add("$fileSize", SqliteType.Integer);
        var parentPathParameter = command.Parameters.Add("$parentPath", SqliteType.Text);

        foreach (var entry in entries)
        {
            cancellationToken.ThrowIfCancellationRequested();

            fileKeyParameter.Value = entry.FileKey;
            projectParameter.Value = entry.ProjectTitle;
            revisionParameter.Value = entry.Revision;
            documentNameParameter.Value = entry.DocumentName;
            lastWriteParameter.Value = entry.LastWriteTime;
            fileSizeParameter.Value = entry.FileSize;
            parentPathParameter.Value = NormalizePath(entry.ParentPath);

            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
    }

    /// <inheritdoc/>
    public async Task DeleteByFileKeysAsync(
        IReadOnlyCollection<string> fileKeys,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        if (fileKeys.Count == 0)
            return;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqliteTransaction)await connection.BeginTransactionAsync(cancellationToken);

        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "DELETE FROM FileIdentityCache WHERE FileKey = $fileKey;";
        var parameter = command.Parameters.Add("$fileKey", SqliteType.Text);

        foreach (var key in fileKeys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            parameter.Value = key;
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
    }

    /// <inheritdoc/>
    public async Task DeleteByNormalizedPathAsync(
        string fullPath,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken);

        var normalized = NormalizePath(fullPath);
        if (string.IsNullOrWhiteSpace(normalized))
            return;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = "DELETE FROM FileIdentityCache WHERE FileKey LIKE $prefix;";
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("$prefix", normalized + "|%");
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

            var parentDirectory = Path.GetDirectoryName(DatabasePath);
            if (string.IsNullOrWhiteSpace(parentDirectory))
                throw new InvalidOperationException("Invalid identity cache database path.");

            Directory.CreateDirectory(parentDirectory);

            await using var connection = CreateConnection();
            await connection.OpenAsync(cancellationToken);

            await ExecuteNonQueryAsync(connection, "PRAGMA journal_mode = WAL;", cancellationToken);
            await ExecuteNonQueryAsync(connection, "PRAGMA synchronous = NORMAL;", cancellationToken);

            const string schemaSql = """
                CREATE TABLE IF NOT EXISTS FileIdentityCache (
                    FileKey TEXT PRIMARY KEY,
                    ProjectTitle TEXT,
                    Revision TEXT,
                    DocumentName TEXT,
                    LastWriteTime INTEGER NOT NULL,
                    FileSize INTEGER NOT NULL,
                    ParentPath TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS IX_FileIdentityCache_ParentPath
                    ON FileIdentityCache (ParentPath);

                CREATE INDEX IF NOT EXISTS idx_parent_path
                    ON FileIdentityCache (ParentPath);

                CREATE INDEX IF NOT EXISTS idx_ParentPath
                    ON FileIdentityCache (ParentPath);

                CREATE INDEX IF NOT EXISTS IX_FileIdentityCache_LastWriteTime
                    ON FileIdentityCache (LastWriteTime);
                """;

            await ExecuteNonQueryAsync(connection, schemaSql, cancellationToken);
            _logger.LogInformation("Identity cache initialized at {Path}", DatabasePath);
            _initialized = true;
        }
        finally
        {
            _initGate.Release();
        }
    }

    private static async Task ExecuteNonQueryAsync(SqliteConnection connection, string commandText, CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = commandText;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private SqliteConnection CreateConnection()
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared
        };

        return new SqliteConnection(builder.ToString());
    }

    private static string ResolveDatabasePath()
    {
        var explicitPath = Environment.GetEnvironmentVariable("FILEEXPLORER_IDENTITY_CACHE_DB");
        if (!string.IsNullOrWhiteSpace(explicitPath))
            return Path.GetFullPath(Environment.ExpandEnvironmentVariables(explicitPath));

        var repoRoot = FindRepositoryRoot();
        if (!string.IsNullOrWhiteSpace(repoRoot))
            return Path.Combine(repoRoot, "Development", DbFileName);

        var fallbackRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FileExplorer",
            "Development");

        return Path.Combine(fallbackRoot, DbFileName);
    }

    private static string? FindRepositoryRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        for (var i = 0; i < 12 && current is not null; i++)
        {
            if (File.Exists(Path.Combine(current.FullName, "FileExplorer.sln")))
                return current.FullName;

            current = current.Parent;
        }

        return null;
    }

    private static string StripLongPathPrefix(string path)
    {
        if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            return @"\\" + path[8..];

        if (path.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            return path[4..];

        return path;
    }
}
