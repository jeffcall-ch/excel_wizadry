// Copyright (c) FileExplorer contributors. All rights reserved.
// AOT Note: COM vtable calls for IContextMenu require [DynamicDependency] attributes.
// The static P/Invoke-only methods (SHGetFileInfo, SHFileOperation, ShellExecute) are AOT-safe.

using System.Diagnostics;
using System.Runtime.InteropServices;
using FileExplorer.Core.Interfaces;
using FileExplorer.Infrastructure.Native;
using Microsoft.Extensions.Logging;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Windows Shell integration service providing context menus, icons, drag-drop, and file launching.
/// </summary>
[System.Diagnostics.CodeAnalysis.DynamicallyAccessedMembers(System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.All)]
public sealed class ShellIntegrationService : IShellIntegrationService
{
    private readonly ILogger<ShellIntegrationService> _logger;

    /// <summary>Initializes a new instance of the <see cref="ShellIntegrationService"/> class.</summary>
    public ShellIntegrationService(ILogger<ShellIntegrationService> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public string GetTypeDescription(string path)
    {
        var shfi = new NativeMethods.SHFILEINFO();
        var flags = NativeMethods.SHGFI_TYPENAME | NativeMethods.SHGFI_USEFILEATTRIBUTES;
        uint attrs = Directory.Exists(path) ? NativeMethods.FILE_ATTRIBUTE_DIRECTORY : NativeMethods.FILE_ATTRIBUTE_NORMAL;

        nint result = NativeMethods.SHGetFileInfo(path, attrs, ref shfi, (uint)Marshal.SizeOf<NativeMethods.SHFILEINFO>(), flags);
        return result != 0 ? shfi.szTypeName : string.Empty;
    }

    /// <inheritdoc/>
    public int GetIconIndex(string path, bool isDirectory)
    {
        var shfi = new NativeMethods.SHFILEINFO();
        var flags = NativeMethods.SHGFI_SYSICONINDEX | NativeMethods.SHGFI_SMALLICON | NativeMethods.SHGFI_USEFILEATTRIBUTES;
        uint attrs = isDirectory ? NativeMethods.FILE_ATTRIBUTE_DIRECTORY : NativeMethods.FILE_ATTRIBUTE_NORMAL;

        NativeMethods.SHGetFileInfo(path, attrs, ref shfi, (uint)Marshal.SizeOf<NativeMethods.SHFILEINFO>(), flags);
        return shfi.iIcon;
    }

    /// <inheritdoc/>
    public nint GetIconHandle(string path, bool smallIcon)
    {
        var shfi = new NativeMethods.SHFILEINFO();
        var flags = NativeMethods.SHGFI_ICON | (smallIcon ? NativeMethods.SHGFI_SMALLICON : NativeMethods.SHGFI_LARGEICON);

        nint result = NativeMethods.SHGetFileInfo(path, 0, ref shfi, (uint)Marshal.SizeOf<NativeMethods.SHFILEINFO>(), flags);
        return result != 0 ? shfi.hIcon : 0;
    }

    /// <inheritdoc/>
    public async Task ShowContextMenuAsync(nint hwnd, IReadOnlyList<string> paths, double x, double y, CancellationToken cancellationToken = default)
    {
        if (paths.Count == 0) return;

        await Task.Run(() =>
        {
            nint hMenu = 0;
            var pidls = new List<nint>();
            NativeMethods.IContextMenu? contextMenu = null;

            try
            {
                // Parse each item to get its PIDL
                var childPidls = new List<nint>();
                NativeMethods.IShellFolder? parentFolder = null;

                if (paths.Count == 1)
                {
                    NativeMethods.SHParseDisplayName(paths[0], 0, out nint pidl, 0, out _);
                    pidls.Add(pidl);

                    NativeMethods.SHBindToParent(pidl, NativeMethods.IID_IShellFolder, out parentFolder, out nint childPidl);
                    childPidls.Add(childPidl);
                }
                else
                {
                    // For multi-select: all items must be in the same parent
                    var parentPath = Path.GetDirectoryName(paths[0]);
                    if (parentPath is null) return;

                    NativeMethods.SHParseDisplayName(parentPath, 0, out nint parentPidl, 0, out _);
                    pidls.Add(parentPidl);

                    NativeMethods.SHGetDesktopFolder(out var desktop);
                    uint eaten;
                    uint attrs = 0;
                    desktop.ParseDisplayName(0, 0, parentPath, out eaten, out nint folderPidl, ref attrs);
                    pidls.Add(folderPidl);
                    desktop.BindToObject(folderPidl, 0, NativeMethods.IID_IShellFolder, out nint ppv);
                    parentFolder = (NativeMethods.IShellFolder)Marshal.GetObjectForIUnknown(ppv);

                    foreach (var path in paths)
                    {
                        var fileName = Path.GetFileName(path);
                        uint e2;
                        uint a2 = 0;
                        parentFolder.ParseDisplayName(0, 0, fileName, out e2, out nint childPidl, ref a2);
                        childPidls.Add(childPidl);
                        pidls.Add(childPidl);
                    }
                }

                // Get IContextMenu
                uint reserved = 0;
                parentFolder!.GetUIObjectOf(hwnd, (uint)childPidls.Count, childPidls.ToArray(), NativeMethods.IID_IContextMenu, ref reserved, out nint ctxPtr);
                contextMenu = (NativeMethods.IContextMenu)Marshal.GetObjectForIUnknown(ctxPtr);

                // Create popup menu and populate
                hMenu = NativeMethods.CreatePopupMenu();
                contextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, NativeMethods.CMF_NORMAL | NativeMethods.CMF_CANRENAME | NativeMethods.CMF_EXPLORE);

                // Show popup
                int cmd = NativeMethods.TrackPopupMenuEx(
                    hMenu,
                    NativeMethods.TPM_RETURNCMD | NativeMethods.TPM_LEFTALIGN,
                    (int)x, (int)y, hwnd, 0);

                if (cmd > 0)
                {
                    var info = new NativeMethods.CMINVOKECOMMANDINFOEX
                    {
                        cbSize = Marshal.SizeOf<NativeMethods.CMINVOKECOMMANDINFOEX>(),
                        fMask = NativeMethods.CMIC_MASK_UNICODE,
                        hwnd = hwnd,
                        lpVerb = (nint)(cmd - 1),
                        lpVerbW = (nint)(cmd - 1),
                        nShow = NativeMethods.SW_SHOWNORMAL,
                    };
                    contextMenu.InvokeCommand(ref info);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to show shell context menu");
            }
            finally
            {
                if (hMenu != 0) NativeMethods.DestroyMenu(hMenu);
                foreach (var pidl in pidls)
                {
                    if (pidl != 0) NativeMethods.CoTaskMemFree(pidl);
                }
                if (contextMenu is not null) Marshal.ReleaseComObject(contextMenu);
            }
        }, cancellationToken);
    }

    /// <inheritdoc/>
    public void ShowProperties(nint hwnd, IReadOnlyList<string> paths)
    {
        if (paths.Count == 0) return;

        foreach (var path in paths)
        {
            try
            {
                var objectType = NativeMethods.SHOP_FILEPATH;
                var objectName = path;

                if (TryGetDriveRoot(path, out var driveRoot))
                {
                    objectType = NativeMethods.SHOP_VOLUMEGUID;
                    objectName = driveRoot;
                }

                if (!NativeMethods.SHObjectProperties(hwnd, objectType, objectName, null))
                {
                    var error = Marshal.GetLastWin32Error();
                    _logger.LogWarning("SHObjectProperties failed for {Path}: {ErrorCode} {ErrorMessage}", path, error, NativeMethods.FormatWin32Error(error));
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to show properties for {Path}", path);
            }
        }
    }

    private static bool TryGetDriveRoot(string path, out string driveRoot)
    {
        driveRoot = string.Empty;
        var root = Path.GetPathRoot(path);
        if (string.IsNullOrWhiteSpace(root))
            return false;

        var trimmedPath = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var trimmedRoot = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (!string.Equals(trimmedPath, trimmedRoot, StringComparison.OrdinalIgnoreCase))
            return false;

        driveRoot = trimmedRoot;
        return true;
    }

    /// <inheritdoc/>
    public void ShowShareDialog(nint hwnd, string filePath)
    {
        // Launch via shell verb 'share'
        try
        {
            NativeMethods.ShellExecute(hwnd, "share", filePath, null, null, NativeMethods.SW_SHOW);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to show share dialog for {Path}", filePath);
        }
    }

    /// <inheritdoc/>
    public string? ResolveShellPath(string shellPath)
    {
        if (string.IsNullOrWhiteSpace(shellPath)) return null;

        if (shellPath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            var knownName = shellPath["shell:".Length..];
            return knownName.ToLowerInvariant() switch
            {
                "desktop" => GetKnownFolderPath(NativeMethods.FOLDERID_Desktop),
                "documents" => GetKnownFolderPath(NativeMethods.FOLDERID_Documents),
                "downloads" => GetKnownFolderPath(NativeMethods.FOLDERID_Downloads),
                "music" => GetKnownFolderPath(NativeMethods.FOLDERID_Music),
                "pictures" => GetKnownFolderPath(NativeMethods.FOLDERID_Pictures),
                "videos" => GetKnownFolderPath(NativeMethods.FOLDERID_Videos),
                "profile" => GetKnownFolderPath(NativeMethods.FOLDERID_Profile),
                _ => null
            };
        }

        return shellPath;
    }

    /// <inheritdoc/>
    public void OpenFile(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to open {Path}", path);
        }
    }

    /// <inheritdoc/>
    public string? ResolveShortcutTarget(string shortcutPath)
    {
        if (!File.Exists(shortcutPath) ||
            !string.Equals(Path.GetExtension(shortcutPath), ".lnk", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var shellType = Type.GetTypeFromProgID("WScript.Shell");
        if (shellType is null)
            return null;

        dynamic? shell = Activator.CreateInstance(shellType);
        if (shell is null)
            return null;

        try
        {
            dynamic shortcut = shell.CreateShortcut(shortcutPath);
            var targetPath = shortcut.TargetPath as string;
            return string.IsNullOrWhiteSpace(targetPath) ? null : targetPath;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to resolve shortcut target: {Path}", shortcutPath);
            return null;
        }
        finally
        {
            Marshal.ReleaseComObject(shell);
        }
    }

    /// <inheritdoc/>
    public void OpenWith(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "rundll32.exe",
                Arguments = $"shell32.dll,OpenAs_RunDLL {path}",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to open 'Open with' for {Path}", path);
        }
    }

    /// <inheritdoc/>
    public void RunAsAdministrator(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true,
                Verb = "runas"
            });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to run as administrator: {Path}", path);
        }
    }

    /// <inheritdoc/>
    public Task CreateDesktopShortcutAsync(string targetPath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            var desktopPath = GetKnownFolderPath(NativeMethods.FOLDERID_Desktop);
            if (desktopPath is null) return;

            var shortcutName = Path.GetFileNameWithoutExtension(targetPath) + " - Shortcut.lnk";
            var shortcutPath = Path.Combine(desktopPath, shortcutName);

            // Use WScript.Shell COM to create the shortcut
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType is null) return;

            dynamic? shell = Activator.CreateInstance(shellType);
            if (shell is null) return;

            try
            {
                dynamic shortcut = shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = targetPath;
                shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath);
                shortcut.Save();
            }
            finally
            {
                Marshal.ReleaseComObject(shell);
            }
        }, cancellationToken);
    }

    private static string? GetKnownFolderPath(Guid folderId)
    {
        int hr = NativeMethods.SHGetKnownFolderPath(folderId, 0, 0, out string path);
        return hr == 0 ? path : null;
    }
}
