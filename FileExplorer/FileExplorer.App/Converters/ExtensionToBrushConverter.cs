// Copyright (c) FileExplorer contributors. All rights reserved.

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using Windows.UI;

namespace FileExplorer.App.Converters;

/// <summary>
/// Converts file extensions to themed foreground brushes.
/// Coloring is applied only in dark mode to preserve readability in light mode.
/// </summary>
public sealed class ExtensionToBrushConverter : IValueConverter
{
    private static readonly SolidColorBrush ExcelBrush = new(Color.FromArgb(255, 129, 199, 132));
    private static readonly SolidColorBrush WordBrush = new(Color.FromArgb(255, 100, 181, 246));
    private static readonly SolidColorBrush PdfBrush = new(Color.FromArgb(255, 229, 115, 115)); // #E57373
    private static readonly SolidColorBrush ArchiveBrush = new(Color.FromArgb(255, 255, 213, 79));
    private static readonly SolidColorBrush TextBrush = new(Color.FromArgb(255, 77, 182, 172));

    /// <inheritdoc/>
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is not string extension || string.IsNullOrWhiteSpace(extension))
            return DependencyProperty.UnsetValue;

        if (string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            return PdfBrush;

        if (!IsDarkThemeActive())
            return DependencyProperty.UnsetValue;

        return extension.ToLowerInvariant() switch
        {
            ".xls" or ".xlsx" or ".xlsm" or ".xlsb" or ".xlt" or ".xltx" or ".xltm" => ExcelBrush,
            ".doc" or ".docx" or ".docm" or ".dot" or ".dotx" or ".dotm" => WordBrush,
            ".zip" or ".7z" => ArchiveBrush,
            ".txt" or ".md" => TextBrush,
            _ => DependencyProperty.UnsetValue
        };
    }

    /// <inheritdoc/>
    public object ConvertBack(object value, Type targetType, object parameter, string language)
    {
        throw new NotSupportedException();
    }

    private static bool IsDarkThemeActive()
    {
        if (App.MainWindow?.Content is FrameworkElement root)
            return root.ActualTheme == ElementTheme.Dark;

        return Application.Current.RequestedTheme == ApplicationTheme.Dark;
    }
}
