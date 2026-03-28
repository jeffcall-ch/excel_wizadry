// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;
using Microsoft.UI.Xaml.Data;

namespace FileExplorer.App.Converters;

/// <summary>
/// Converts a <see cref="CloudFileStatus"/> to a descriptive tooltip string.
/// </summary>
public sealed class CloudStatusToToolTipConverter : IValueConverter
{
    /// <inheritdoc/>
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is CloudFileStatus status)
        {
            return status switch
            {
                CloudFileStatus.CloudOnly => "Available online only",
                CloudFileStatus.LocallyAvailable => "Available on this device",
                CloudFileStatus.Syncing => "Syncing...",
                CloudFileStatus.SyncError => "Sync error",
                CloudFileStatus.SharedWithMe => "Shared with you",
                CloudFileStatus.Pending => "Checking status...",
                _ => string.Empty
            };
        }
        return string.Empty;
    }

    /// <inheritdoc/>
    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotSupportedException();
    }
}
