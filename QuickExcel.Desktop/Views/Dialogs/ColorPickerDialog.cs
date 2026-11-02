using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace QuickExcel.Desktop;

/// <summary>A small WPF-only RGB picker so the palette does not depend on WinForms dialogs.</summary>
internal sealed class ColorPickerDialog : Window
{
    private readonly Slider _red = new() { Minimum = 0, Maximum = 255, TickFrequency = 1, IsSnapToTickEnabled = true };
    private readonly Slider _green = new() { Minimum = 0, Maximum = 255, TickFrequency = 1, IsSnapToTickEnabled = true };
    private readonly Slider _blue = new() { Minimum = 0, Maximum = 255, TickFrequency = 1, IsSnapToTickEnabled = true };
    private readonly TextBox _hex = new() { Width = 110, FontFamily = new FontFamily("Consolas") };
    private readonly Border _preview = new() { Height = 52, BorderBrush = Brushes.Gray, BorderThickness = new(1), CornerRadius = new(3) };
    private bool _updating;

    internal string SelectedColor { get; private set; } = "FF000000";

    public ColorPickerDialog(Window owner, string title, string? initial)
    {
        Owner = owner; Title = title; SizeToContent = SizeToContent.WidthAndHeight; ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var (r, g, b) = Parse(initial);
        var panel = new StackPanel { Margin = new(20), Width = 340 };
        panel.Children.Add(new TextBlock { Text = "拖动滑块或输入 RRGGBB 十六进制颜色值。", TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 10) });
        panel.Children.Add(_preview);
        panel.Children.Add(new TextBlock { Text = "红 (R)", Margin = new(0, 10, 0, 2) }); panel.Children.Add(_red);
        panel.Children.Add(new TextBlock { Text = "绿 (G)", Margin = new(0, 6, 0, 2) }); panel.Children.Add(_green);
        panel.Children.Add(new TextBlock { Text = "蓝 (B)", Margin = new(0, 6, 0, 2) }); panel.Children.Add(_blue);
        var hexRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new(0, 10, 0, 0) };
        hexRow.Children.Add(new TextBlock { Text = "颜色值", VerticalAlignment = VerticalAlignment.Center, Margin = new(0, 0, 8, 0) });
        hexRow.Children.Add(_hex); panel.Children.Add(hexRow);
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap, Margin = new(0, 6, 0, 0) }; panel.Children.Add(error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 14, 0, 0) };
        var ok = new Button { Content = "确定", IsDefault = true, MinWidth = 80 }; var cancel = new Button { Content = "取消", IsCancel = true, MinWidth = 80, Margin = new(8, 0, 0, 0) };
        buttons.Children.Add(ok); buttons.Children.Add(cancel); panel.Children.Add(buttons);
        _red.Value = r; _green.Value = g; _blue.Value = b; Content = panel;
        _updating = true; _hex.Text = $"{r:X2}{g:X2}{b:X2}"; _updating = false; RefreshPreview();
        void FromSliders() { _updating = true; _hex.Text = $"{(int)_red.Value:X2}{(int)_green.Value:X2}{(int)_blue.Value:X2}"; _updating = false; RefreshPreview(); }
        _red.ValueChanged += (_, _) => { if (!_updating) FromSliders(); };
        _green.ValueChanged += (_, _) => { if (!_updating) FromSliders(); };
        _blue.ValueChanged += (_, _) => { if (!_updating) FromSliders(); };
        _hex.TextChanged += (_, _) =>
        {
            if (_updating) return;
            if (!TryParseHex(_hex.Text, out var color)) { error.Text = "请输入 6 位 RRGGBB 十六进制颜色。"; return; }
            error.Text = ""; _updating = true; _red.Value = color.R; _green.Value = color.G; _blue.Value = color.B; _updating = false; RefreshPreview();
        };
        ok.Click += (_, _) =>
        {
            if (!TryParseHex(_hex.Text, out var color)) { error.Text = "请输入 6 位 RRGGBB 十六进制颜色。"; return; }
            SelectedColor = $"FF{color.R:X2}{color.G:X2}{color.B:X2}"; DialogResult = true;
        };
        Loaded += (_, _) => { _hex.Focus(); _hex.SelectAll(); };
    }

    private void RefreshPreview()
    {
        var color = Color.FromRgb((byte)_red.Value, (byte)_green.Value, (byte)_blue.Value);
        _preview.Background = new SolidColorBrush(color);
    }

    private static (byte R, byte G, byte B) Parse(string? text)
    {
        if (!TryParseHex(text, out var color)) return (0, 0, 0);
        return (color.R, color.G, color.B);
    }

    private static bool TryParseHex(string? text, out Color color)
    {
        color = Colors.Black;
        if (text is null) return false;
        var value = text.Trim().TrimStart('#');
        if (value.Length == 8) value = value[2..];
        if (value.Length != 6 || !uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)) return false;
        color = Color.FromRgb((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
        return true;
    }
}
