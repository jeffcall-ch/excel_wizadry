// Copyright (c) FileExplorer contributors. All rights reserved.

namespace FileExplorer.Core.Models;

/// <summary>
/// Represents the OneDrive cloud synchronization status of a file.
/// </summary>
public enum CloudFileStatus
{
    /// <summary>File is not under any cloud sync root.</summary>
    NotApplicable = 0,

    /// <summary>File exists only in the cloud and is not downloaded locally.</summary>
    CloudOnly = 1,

    /// <summary>File is fully available locally.</summary>
    LocallyAvailable = 2,

    /// <summary>File is pinned to stay always available on this device.</summary>
    AlwaysAvailable = 7,

    /// <summary>File is currently syncing.</summary>
    Syncing = 3,

    /// <summary>Sync error has occurred for this file.</summary>
    SyncError = 4,

    /// <summary>File has been shared with the current user.</summary>
    SharedWithMe = 5,

    /// <summary>Cloud status is being queried.</summary>
    Pending = 6
}
