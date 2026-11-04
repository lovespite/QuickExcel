using Avalonia.Controls;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

internal static class ColorPalette
{
    public static readonly (string Name, string Value)[] Common =
    [
        ("黑色", "FF000000"), ("白色", "FFFFFFFF"), ("灰色", "FF6B7280"), ("红色", "FFFF0000"),
        ("橙色", "FFFFA500"), ("黄色", "FFFFFF00"), ("绿色", "FF008000"), ("青色", "FF0D9488"),
        ("蓝色", "FF0000FF"), ("紫色", "FF9333EA"), ("粉色", "FFDB2777"), ("浅蓝色", "FFBAE6FD")
    ];

    public static bool TryParse(string? text, out Color color)
    {
        var value = text?.Trim() ?? "";
        if (value.StartsWith('#')) value = value[1..];
        color = default;
        return value.Length is 6 or 8 && value.All(Uri.IsHexDigit) && Color.TryParse("#" + value, out color);
    }

    public static Border Swatch(string? value, double size = 14) => new()
    {
        Width = size, Height = size, CornerRadius = new(2), BorderBrush = Brushes.Gray, BorderThickness = new(1),
        Background = TryParse(value, out var color) ? new SolidColorBrush(color) : Brushes.Transparent,
        Child = value is null ? new TextBlock { Text = "╱", FontSize = 11, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Center } : null
    };
}
