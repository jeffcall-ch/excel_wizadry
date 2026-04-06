// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Infrastructure.Services;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace FileExplorer.Tests;

public sealed class SqliteIdentityCacheServiceTests : IDisposable
{
    private readonly string _tempRoot;
    private readonly string _dbPath;
    private readonly string? _previousOverride;

    public SqliteIdentityCacheServiceTests()
    {
        _tempRoot = Path.Combine(Path.GetTempPath(), "identity-cache-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempRoot);

        _dbPath = Path.Combine(_tempRoot, "identity_cache.db");
        _previousOverride = Environment.GetEnvironmentVariable("FILEEXPLORER_IDENTITY_CACHE_DB");
        Environment.SetEnvironmentVariable("FILEEXPLORER_IDENTITY_CACHE_DB", _dbPath);
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("FILEEXPLORER_IDENTITY_CACHE_DB", _previousOverride);

        try
        {
            if (Directory.Exists(_tempRoot))
                Directory.Delete(_tempRoot, recursive: true);
        }
        catch
        {
            // Best effort cleanup for temporary test artifacts.
        }
    }

    [Fact]
    public async Task UpsertAndFetchByParentPath_Works()
    {
        var service = new SqliteIdentityCacheService(NullLogger<SqliteIdentityCacheService>.Instance);

        var parent = service.NormalizePath(@"C:\Work\Folder");
        var key = service.BuildFileKey(@"C:\Work\Folder\Doc.xlsx", 123, new DateTimeOffset(2026, 4, 6, 12, 0, 0, TimeSpan.Zero));

        await service.UpsertEntriesAsync([
            new FileExplorer.Core.Models.FileIdentityCacheRecord
            {
                FileKey = key,
                ParentPath = parent,
                ProjectTitle = "Project A",
                Revision = "1.0",
                DocumentName = "Doc",
                LastWriteTime = 638795088000000000,
                FileSize = 123
            }
        ]);

        var rows = await service.GetEntriesByParentPathAsync(parent);

        rows.Should().ContainKey(key);
        rows[key].ProjectTitle.Should().Be("Project A");
        rows[key].Revision.Should().Be("1.0");
    }

    [Fact]
    public async Task DeleteByNormalizedPath_RemovesAllVersions()
    {
        var service = new SqliteIdentityCacheService(NullLogger<SqliteIdentityCacheService>.Instance);

        var parent = service.NormalizePath(@"C:\Work\Folder");
        var path = @"C:\Work\Folder\Doc.xlsx";

        var keyV1 = service.BuildFileKey(path, 100, new DateTimeOffset(2026, 4, 6, 12, 0, 0, TimeSpan.Zero));
        var keyV2 = service.BuildFileKey(path, 120, new DateTimeOffset(2026, 4, 6, 12, 5, 0, TimeSpan.Zero));

        await service.UpsertEntriesAsync([
            new FileExplorer.Core.Models.FileIdentityCacheRecord
            {
                FileKey = keyV1,
                ParentPath = parent,
                ProjectTitle = "P1",
                Revision = "1.0",
                DocumentName = "Doc",
                LastWriteTime = 638795088000000000,
                FileSize = 100
            },
            new FileExplorer.Core.Models.FileIdentityCacheRecord
            {
                FileKey = keyV2,
                ParentPath = parent,
                ProjectTitle = "P2",
                Revision = "2.0",
                DocumentName = "Doc",
                LastWriteTime = 638795091000000000,
                FileSize = 120
            }
        ]);

        await service.DeleteByNormalizedPathAsync(path);

        var rows = await service.GetEntriesByParentPathAsync(parent);
        rows.Should().NotContainKey(keyV1);
        rows.Should().NotContainKey(keyV2);
    }
}
