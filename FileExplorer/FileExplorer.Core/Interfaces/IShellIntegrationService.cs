// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;

namespace FileExplorer.Core.Interfaces;

/// <summary>
/// Service for Windows Shell integration including context menus, icons, and drag-drop.
/// </summary>
public interface IShellIntegrationService
{
    /// <summary>Gets the shell type description for a file (e.g., "Text Document").</summary>
    string GetTypeDescription(string path);

    /// <summary>Gets the shell icon index for a file or directory.</summary>
    int GetIconIndex(string path, bool isDirectory);

    /// <summary>Gets the icon image source for the specified icon index.</summary>
    nint GetIconHandle(string path, bool smallIcon);

    /// <summary>Shows the native shell context menu at the specified screen position.</summary>
    Task ShowContextMenuAsync(nint hwnd, IReadOnlyList<string> paths, double x, double y, CancellationToken cancellationToken = default);

    /// <summary>Shows the Properties dialog for the specified paths.</summary>
    void ShowProperties(nint hwnd, IReadOnlyList<string> paths);

    /// <summary>Opens the Windows Share dialog for the specified file.</summary>
    void ShowShareDialog(nint hwnd, string filePath);

    /// <summary>Resolves a shell path (e.g., "shell:Downloads") to a filesystem path.</summary>
    string? ResolveShellPath(string shellPath);

    /// <summary>Resolves a Windows shortcut (.lnk) to its target path.</summary>
    string? ResolveShortcutTarget(string shortcutPath);

    /// <summary>Opens a file with the default handler.</summary>
    void OpenFile(string path);

    /// <summary>Opens "Open with" dialog for a file.</summary>
    void OpenWith(string path);

    /// <summary>Runs a file as administrator.</summary>
    void RunAsAdministrator(string path);

    /// <summary>Creates a shortcut on the desktop for the given path.</summary>
    Task CreateDesktopShortcutAsync(string targetPath, CancellationToken cancellationToken = default);
}
