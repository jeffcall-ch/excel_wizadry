// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Models;
using Microsoft.UI.Xaml.Data;

namespace FileExplorer.App.Converters;

/// <summary>
/// Converts a <see cref="CloudFileStatus"/> to a Segoe Fluent Icons glyph string.
/// </summary>
public sealed class CloudStatusToIconConverter : IValueConverter
{
    /// <inheritdoc/>
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is CloudFileStatus status)
        {
            return status switch
            {
                CloudFileStatus.CloudOnly => "\uE753",        // Cloud
                CloudFileStatus.LocallyAvailable => "\uE73E", // Checkmark
                CloudFileStatus.AlwaysAvailable => "\uE718",  // Pinned
                CloudFileStatus.Syncing => "\uE895",          // Sync
                CloudFileStatus.SyncError => "\uE7BA",        // Warning
                CloudFileStatus.SharedWithMe => "\uE902",     // People
                CloudFileStatus.Pending => "\uE712",          // More
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
