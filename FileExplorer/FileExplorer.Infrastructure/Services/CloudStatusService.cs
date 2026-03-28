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

            // Check common OneDrive locations
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var oneDriveDir = Path.Combine(localAppData, "Microsoft", "OneDrive", "settings");

            // Also check user profile for OneDrive directories
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var potentialRoots = new[]
            {
                Path.Combine(userProfile, "OneDrive"),
                Path.Combine(userProfile, "OneDrive - Personal"),
            };

            foreach (var root in potentialRoots)
            {
                if (Directory.Exists(root))
                {
                    try
                    {
                        // Verify it's actually a sync root via CloudFilter API
                        int bufferSize = 4096;
                        nint buffer = Marshal.AllocHGlobal(bufferSize);
                        try
                        {
                            int hr = NativeMethods.CfGetSyncRootInfoByPath(root, 0, buffer, bufferSize, out _);
                            if (hr == 0)
                            {
                                _syncRoots.Add(root);
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
            }

            // Search for corporate OneDrive folders
            try
            {
                foreach (var dir in Directory.EnumerateDirectories(userProfile, "OneDrive*"))
                {
                    if (!_syncRoots.Contains(dir) && Directory.Exists(dir))
                    {
                        try
                        {
                            int bufferSize = 4096;
                            nint buffer = Marshal.AllocHGlobal(bufferSize);
                            try
                            {
                                int hr = NativeMethods.CfGetSyncRootInfoByPath(dir, 0, buffer, bufferSize, out _);
                                if (hr == 0)
                                {
                                    _syncRoots.Add(dir);
                                    _logger.LogInformation("Detected OneDrive sync root: {Path}", dir);
                                }
                            }
                            finally
                            {
                                Marshal.FreeHGlobal(buffer);
                            }
                        }
                        catch { /* Not a sync root */ }
                    }
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
            filePath.StartsWith(root, StringComparison.OrdinalIgnoreCase));
    }

    /// <inheritdoc/>
    public Task<CloudFileStatus> GetCloudStatusAsync(string filePath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            if (!IsUnderSyncRoot(filePath)) return CloudFileStatus.NotApplicable;

            try
            {
                var fileInfo = new FileInfo(filePath);
                if (!fileInfo.Exists) return CloudFileStatus.NotApplicable;

                var attrs = fileInfo.Attributes;

                // FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS (0x00400000) = cloud-only placeholder
                // FILE_ATTRIBUTE_RECALL_ON_OPEN (0x00040000) = partial
                // FILE_ATTRIBUTE_PINNED (0x00080000) = locally pinned
                // FILE_ATTRIBUTE_UNPINNED (0x00100000) = free up space candidate

                const FileAttributes recallOnDataAccess = (FileAttributes)0x00400000;
                const FileAttributes recallOnOpen = (FileAttributes)0x00040000;
                const FileAttributes pinned = (FileAttributes)0x00080000;
                const FileAttributes offline = FileAttributes.Offline;

                if ((attrs & recallOnDataAccess) != 0 || (attrs & offline) != 0)
                    return CloudFileStatus.CloudOnly;

                if ((attrs & recallOnOpen) != 0)
                    return CloudFileStatus.Syncing;

                if ((attrs & pinned) != 0)
                    return CloudFileStatus.LocallyAvailable;

                // File exists locally under sync root — locally available
                return CloudFileStatus.LocallyAvailable;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed to get cloud status for {Path}", filePath);
                return CloudFileStatus.NotApplicable;
            }
        }, cancellationToken);
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
}
