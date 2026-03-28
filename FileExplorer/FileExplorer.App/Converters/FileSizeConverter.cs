// Copyright (c) FileExplorer contributors. All rights reserved.

using Microsoft.UI.Xaml.Data;

namespace FileExplorer.App.Converters;

/// <summary>
/// Converts a byte count to a human-readable file size string.
/// </summary>
public sealed class FileSizeConverter : IValueConverter
{
    /// <inheritdoc/>
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is long bytes)
        {
            return bytes switch
            {
                0 => string.Empty,
                < 1024 => $"{bytes} B",
                < 1024 * 1024 => $"{bytes / 1024.0:F1} KB",
                < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):F1} MB",
                _ => $"{bytes / (1024.0 * 1024 * 1024):F2} GB"
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
