using Avalonia.Controls;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

internal static class ColorPalette
{
    public static readonly (string Name, string Value)[] Common =
    [
        (UiText.Get("黑色"), "FF000000"), (UiText.Get("白色"), "FFFFFFFF"), (UiText.Get("灰色"), "FF6B7280"), (UiText.Get("红色"), "FFFF0000"),
        (UiText.Get("橙色"), "FFFFA500"), (UiText.Get("黄色"), "FFFFFF00"), (UiText.Get("绿色"), "FF008000"), (UiText.Get("青色"), "FF0D9488"),
        (UiText.Get("蓝色"), "FF0000FF"), (UiText.Get("紫色"), "FF9333EA"), (UiText.Get("粉色"), "FFDB2777"), (UiText.Get("浅蓝色"), "FFBAE6FD")
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
