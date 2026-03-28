// Copyright (c) FileExplorer contributors. All rights reserved.

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;

namespace FileExplorer.App.Converters;

/// <summary>
/// Converts a boolean value to a <see cref="Visibility"/> value.
/// </summary>
public sealed class BoolToVisibilityConverter : IValueConverter
{
    /// <summary>When true, inverts the conversion (true = Collapsed, false = Visible).</summary>
    public bool IsInverted { get; set; }

    /// <inheritdoc/>
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        bool boolValue = value is bool b && b;
        if (IsInverted) boolValue = !boolValue;
        return boolValue ? Visibility.Visible : Visibility.Collapsed;
    }

    /// <inheritdoc/>
    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        bool visible = value is Visibility v && v == Visibility.Visible;
        return IsInverted ? !visible : visible;
    }
}
