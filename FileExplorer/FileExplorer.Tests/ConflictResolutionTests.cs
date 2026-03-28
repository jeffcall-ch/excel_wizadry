// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;
using Xunit;
using FluentAssertions;

namespace FileExplorer.Tests;

/// <summary>
/// Unit tests for <see cref="FileConflict"/> resolution logic.
/// </summary>
public class ConflictResolutionTests
{
    [Fact]
    public void ConflictResolutionResult_Skip_HasCorrectValues()
    {
        var result = new ConflictResolutionResult
        {
            Resolution = ConflictResolution.Skip,
            ApplyToAll = false
        };

        result.Resolution.Should().Be(ConflictResolution.Skip);
        result.ApplyToAll.Should().BeFalse();
    }

    [Fact]
    public void ConflictResolutionResult_ReplaceAll_HasCorrectValues()
    {
        var result = new ConflictResolutionResult
        {
            Resolution = ConflictResolution.Replace,
            ApplyToAll = true
        };

        result.Resolution.Should().Be(ConflictResolution.Replace);
        result.ApplyToAll.Should().BeTrue();
    }

    [Fact]
    public void FileConflict_CapturesMetadata()
    {
        var conflict = new FileConflict
        {
            SourcePath = @"C:\source\file.txt",
            DestinationPath = @"C:\dest\file.txt",
            SourceSize = 1024,
            DestinationSize = 2048,
            SourceDate = new DateTimeOffset(2025, 1, 1, 0, 0, 0, TimeSpan.Zero),
            DestinationDate = new DateTimeOffset(2024, 6, 15, 0, 0, 0, TimeSpan.Zero)
        };

        conflict.SourcePath.Should().Be(@"C:\source\file.txt");
        conflict.DestinationPath.Should().Be(@"C:\dest\file.txt");
        conflict.SourceSize.Should().Be(1024);
        conflict.DestinationSize.Should().Be(2048);
        conflict.SourceDate.Should().BeAfter(conflict.DestinationDate);
    }

    [Fact]
    public void FileOperation_HasUniqueId()
    {
        var op1 = new FileOperation
        {
            OperationType = FileOperationType.Copy,
            SourcePaths = ["path1"]
        };

        var op2 = new FileOperation
        {
            OperationType = FileOperationType.Move,
            SourcePaths = ["path2"]
        };

        op1.Id.Should().NotBe(op2.Id);
        op1.Id.Should().NotBe(Guid.Empty);
    }

    [Fact]
    public void FileOperationProgress_CalculatesPercentCorrectly()
    {
        var progress = new FileOperationProgress
        {
            OperationId = Guid.NewGuid(),
            TotalBytes = 1000,
            BytesCompleted = 500
        };

        progress.ProgressPercent.Should().BeApproximately(50.0, 0.01);
    }

    [Fact]
    public void FileOperationProgress_ZeroTotalBytes_ReturnsZeroPercent()
    {
        var progress = new FileOperationProgress
        {
            OperationId = Guid.NewGuid(),
            TotalBytes = 0,
            BytesCompleted = 0
        };

        progress.ProgressPercent.Should().Be(0);
    }

    [Fact]
    public void UndoableOperation_CapturesAllFields()
    {
        var op = new UndoableOperation
        {
            OriginalOperation = FileOperationType.Rename,
            Description = "Rename file.txt to newfile.txt",
            OriginalSourcePaths = [@"C:\file.txt"],
            ResultPaths = [@"C:\newfile.txt"],
            PreviousName = "file.txt"
        };

        op.OriginalOperation.Should().Be(FileOperationType.Rename);
        op.Description.Should().Contain("Rename");
        op.PreviousName.Should().Be("file.txt");
        op.Id.Should().NotBe(Guid.Empty);
        op.Timestamp.Should().BeCloseTo(DateTimeOffset.Now, TimeSpan.FromSeconds(5));
    }
}
