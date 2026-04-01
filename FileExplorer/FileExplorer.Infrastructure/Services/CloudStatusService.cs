// Copyright (c) FileExplorer contributors. All rights reserved.
// AOT Note: Cloud Filter API calls are plain P/Invoke — no COM or reflection involved. AOT-safe.

using System.Runtime.InteropServices;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Infrastructure.Native;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Cloud Filter API integration for OneDrive placeholder state querying and pin/dehydrate operations.
/// </summary>
public sealed class CloudStatusService : ICloudStatusService
{
    private readonly ILogger<CloudStatusService> _logger;
    private readonly List<string> _syncRoots = [];
    private readonly SemaphoreSlim _batchSemaphore = new(1, 1);

    /// <inheritdoc/>
    public bool IsSyncRootDetected => _syncRoots.Count > 0;

    /// <summary>Initializes a new instance of the <see cref="CloudStatusService"/> class.</summary>
    public CloudStatusService(ILogger<CloudStatusService> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public Task DetectSyncRootsAsync(CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            _syncRoots.Clear();

            // Also check user profile for OneDrive directories
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var potentialRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                Path.Combine(userProfile, "OneDrive"),
                Path.Combine(userProfile, "OneDrive - Personal"),
            };

            foreach (var envVar in new[] { "OneDrive", "OneDriveConsumer", "OneDriveCommercial" })
            {
                var envPath = Environment.GetEnvironmentVariable(envVar);
                if (!string.IsNullOrWhiteSpace(envPath))
                {
                    potentialRoots.Add(envPath);
                }
            }

            foreach (var root in potentialRoots)
            {
                TryAddSyncRoot(root);
            }

            // Search for corporate OneDrive folders
            try
            {
                foreach (var dir in Directory.EnumerateDirectories(userProfile, "OneDrive*"))
                {
                    TryAddSyncRoot(dir);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Error searching for OneDrive directories");
            }

            _logger.LogInformation("Detected {Count} OneDrive sync root(s)", _syncRoots.Count);
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public bool IsUnderSyncRoot(string filePath)
    {
        return _syncRoots.Any(root =>
            filePath.Equals(root, StringComparison.OrdinalIgnoreCase) ||
            filePath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
            filePath.StartsWith(root + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase));
    }

    /// <inheritdoc/>
    public Task<CloudFileStatus> GetCloudStatusAsync(string filePath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            if (!IsUnderSyncRoot(filePath)) return CloudFileStatus.NotApplicable;

            try
            {
                if (!NativeMethods.GetFileAttributesEx(filePath, 0, out var attributeData))
                    return CloudFileStatus.NotApplicable;

                var attributes = (uint)attributeData.dwFileAttributes;
                uint reparseTag = 0;

                if ((attributes & NativeMethods.FILE_ATTRIBUTE_REPARSE_POINT) != 0)
                {
                    var isDirectory = (attributeData.dwFileAttributes & FileAttributes.Directory) != 0;
                    TryGetReparseTag(filePath, isDirectory, out reparseTag);
                }

                return GetCloudStatus(attributes, reparseTag);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed to get cloud status for {Path}", filePath);
                return CloudFileStatus.NotApplicable;
            }
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public CloudFileStatus GetCloudStatus(uint fileAttributes, uint reparseTag)
    {
        if ((fileAttributes & NativeMethods.FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) != 0)
            return CloudFileStatus.CloudOnly;

        if ((fileAttributes & NativeMethods.FILE_ATTRIBUTE_PINNED) != 0)
            return CloudFileStatus.AlwaysAvailable;

        if ((fileAttributes & NativeMethods.FILE_ATTRIBUTE_REPARSE_POINT) != 0 && IsCloudReparseTag(reparseTag))
        {
            var placeholderState = NativeMethods.CfGetPlaceholderStateFromAttributeTag(fileAttributes, reparseTag);

            if ((placeholderState & NativeMethods.CF_PLACEHOLDER_STATE.CF_PLACEHOLDER_STATE_PARTIAL) != 0)
                return CloudFileStatus.Syncing;

            if ((placeholderState & NativeMethods.CF_PLACEHOLDER_STATE.CF_PLACEHOLDER_STATE_IN_SYNC) != 0 ||
                (placeholderState & NativeMethods.CF_PLACEHOLDER_STATE.CF_PLACEHOLDER_STATE_PARTIALLY_IN_SYNC) != 0)
            {
                return CloudFileStatus.LocallyAvailable;
            }
        }

        if ((fileAttributes & NativeMethods.FILE_ATTRIBUTE_OFFLINE) != 0)
            return CloudFileStatus.CloudOnly;

        return CloudFileStatus.NotApplicable;
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyDictionary<string, CloudFileStatus>> GetCloudStatusBatchAsync(
        IReadOnlyList<string> filePaths, CancellationToken cancellationToken = default)
    {
        await _batchSemaphore.WaitAsync(cancellationToken);
        try
        {
            var results = new Dictionary<string, CloudFileStatus>(filePaths.Count);
            var tasks = new List<Task<(string Path, CloudFileStatus Status)>>();

            foreach (var path in filePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();
                tasks.Add(Task.Run(async () =>
                {
                    var status = await GetCloudStatusAsync(path, cancellationToken);
                    return (path, status);
                }, cancellationToken));
            }

            var completed = await Task.WhenAll(tasks);
            foreach (var (path, status) in completed)
            {
                results[path] = status;
            }

            return results;
        }
        finally
        {
            _batchSemaphore.Release();
        }
    }

    /// <inheritdoc/>
    public Task<IReadOnlyDictionary<string, CloudFileStatus>> GetCloudStatusBatchAsync(
        IReadOnlyList<FileSystemEntry> entries,
        CancellationToken cancellationToken = default)
    {
        return Task.Run<IReadOnlyDictionary<string, CloudFileStatus>>(() =>
        {
            var results = new Dictionary<string, CloudFileStatus>(entries.Count, StringComparer.OrdinalIgnoreCase);

            foreach (var entry in entries)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!IsUnderSyncRoot(entry.FullPath))
                {
                    results[entry.FullPath] = CloudFileStatus.NotApplicable;
                    continue;
                }

                results[entry.FullPath] = GetCloudStatus((uint)entry.Attributes, entry.ReparseTag);
            }

            return results;
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public Task PinFileAsync(string filePath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            int hr = NativeMethods.CfOpenFileWithOplock(filePath, 0, out nint handle);
            if (hr != 0)
            {
                _logger.LogWarning("CfOpenFileWithOplock failed for {Path}: HR=0x{Hr:X8}", filePath, hr);
                return;
            }

            try
            {
                hr = NativeMethods.CfSetPinState(handle, NativeMethods.CF_PIN_STATE_PINNED, 0, 0);
                if (hr != 0)
                {
                    _logger.LogWarning("CfSetPinState (Pin) failed for {Path}: HR=0x{Hr:X8}", filePath, hr);
                }

                hr = NativeMethods.CfHydratePlaceholder(handle, 0, -1, 0, 0);
                if (hr != 0)
                {
                    _logger.LogWarning("CfHydratePlaceholder failed for {Path}: HR=0x{Hr:X8}", filePath, hr);
                }
            }
            finally
            {
                NativeMethods.CfCloseHandle(handle);
            }
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public Task DehydrateFileAsync(string filePath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            int hr = NativeMethods.CfOpenFileWithOplock(filePath, 0, out nint handle);
            if (hr != 0)
            {
                _logger.LogWarning("CfOpenFileWithOplock failed for {Path}: HR=0x{Hr:X8}", filePath, hr);
                return;
            }

            try
            {
                hr = NativeMethods.CfSetPinState(handle, NativeMethods.CF_PIN_STATE_UNPINNED, 0, 0);
                if (hr != 0)
                {
                    _logger.LogWarning("CfSetPinState (Unpin) failed for {Path}: HR=0x{Hr:X8}", filePath, hr);
                }

                hr = NativeMethods.CfDehydratePlaceholder(handle, 0, -1, 0, 0);
                if (hr != 0)
                {
                    _logger.LogWarning("CfDehydratePlaceholder failed for {Path}: HR=0x{Hr:X8}", filePath, hr);
                }
            }
            finally
            {
                NativeMethods.CfCloseHandle(handle);
            }
        }, cancellationToken);
    }

    private void TryAddSyncRoot(string root)
    {
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root) || _syncRoots.Contains(root))
            return;

        try
        {
            int bufferSize = 4096;
            nint buffer = Marshal.AllocHGlobal(bufferSize);
            try
            {
                int hr = NativeMethods.CfGetSyncRootInfoByPath(root, 0, buffer, bufferSize, out _);
                if (hr == 0)
                {
                    _syncRoots.Add(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                    _logger.LogInformation("Detected OneDrive sync root: {Path}", root);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Path {Path} is not a cloud sync root", root);
        }
    }

    private bool TryGetReparseTag(string path, bool isDirectory, out uint reparseTag)
    {
        reparseTag = 0;

        var flags = isDirectory ? NativeMethods.FILE_FLAG_BACKUP_SEMANTICS : 0u;
        var handle = NativeMethods.CreateFile(
            path,
            NativeMethods.FILE_READ_ATTRIBUTES,
            NativeMethods.FILE_SHARE_READ | NativeMethods.FILE_SHARE_WRITE | NativeMethods.FILE_SHARE_DELETE,
            0,
            NativeMethods.OPEN_EXISTING,
            flags,
            0);

        if (handle == NativeMethods.INVALID_HANDLE_VALUE)
        {
            return false;
        }

        try
        {
            var info = new NativeMethods.FILE_ATTRIBUTE_TAG_INFO();
            var bufferSize = Marshal.SizeOf<NativeMethods.FILE_ATTRIBUTE_TAG_INFO>();
            var buffer = Marshal.AllocHGlobal(bufferSize);

            try
            {
                Marshal.StructureToPtr(info, buffer, false);
                if (!NativeMethods.GetFileInformationByHandleEx(handle, NativeMethods.FileAttributeTagInfo, buffer, (uint)bufferSize))
                {
                    return false;
                }

                var tagInfo = Marshal.PtrToStructure<NativeMethods.FILE_ATTRIBUTE_TAG_INFO>(buffer);
                reparseTag = tagInfo.ReparseTag;
                return true;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
        finally
        {
            NativeMethods.CloseHandle(handle);
        }
    }

    private static bool IsCloudReparseTag(uint tag)
    {
        return tag == NativeMethods.IO_REPARSE_TAG_CLOUD ||
               tag == NativeMethods.IO_REPARSE_TAG_CLOUD_1 ||
               tag == NativeMethods.IO_REPARSE_TAG_CLOUD_2 ||
               tag == NativeMethods.IO_REPARSE_TAG_CLOUD_3;
    }
}
