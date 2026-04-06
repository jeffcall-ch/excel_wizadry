// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Infrastructure.Services;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace FileExplorer.Tests;

public sealed class RqUtilityServiceTests : IDisposable
{
    private readonly string _testDir;

    public RqUtilityServiceTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), "RqUtilityTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_testDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
                Directory.Delete(_testDir, recursive: true);
        }
        catch
        {
            // Best-effort cleanup for temp test assets.
        }
    }

    [Fact]
    public async Task GetFolderAggregateAsync_NonexistentFolder_ReturnsFailure()
    {
        var service = new RqUtilityService(NullLogger<RqUtilityService>.Instance);

        var result = await service.GetFolderAggregateAsync(Path.Combine(_testDir, "missing"));

        result.IsSuccess.Should().BeFalse();
        result.ErrorMessage.Should().Contain("Folder not found");
    }

    [Fact]
    public async Task GetFolderAggregateAsync_ComputesRecursiveCountAndSize()
    {
        var service = new RqUtilityService(NullLogger<RqUtilityService>.Instance);

        var nestedDir = Path.Combine(_testDir, "nested");
        Directory.CreateDirectory(nestedDir);

        var fileA = Path.Combine(_testDir, "a.txt");
        var fileB = Path.Combine(nestedDir, "b.txt");

        await File.WriteAllTextAsync(fileA, new string('A', 10));
        await File.WriteAllTextAsync(fileB, new string('B', 30));

        var result = await service.GetFolderAggregateAsync(_testDir);

        result.IsSuccess.Should().BeTrue();
        result.FileCount.Should().Be(2);
        result.TotalSizeBytes.Should().Be(40);
        result.ElapsedMilliseconds.Should().BeGreaterThanOrEqualTo(0);
    }

    [Fact]
    public async Task GetFolderAggregateAsync_PreCanceledToken_ThrowsOperationCanceledException()
    {
        var service = new RqUtilityService(NullLogger<RqUtilityService>.Instance);

        using var cts = new CancellationTokenSource();
        cts.Cancel();

        Func<Task> act = async () => await service.GetFolderAggregateAsync(_testDir, cts.Token);

        await act.Should().ThrowAsync<OperationCanceledException>();
    }
}
