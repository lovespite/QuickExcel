using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Reusable color editor. Only confirmation returns a color; previews never edit a document.</summary>
internal sealed class CustomColorDialog : Window
{
    private readonly TextBox _hex = new() { Name = "ColorHex", MinHeight = 32 };
    private readonly Border _preview = new() { Height = 52, CornerRadius = new(4), BorderThickness = new(1), BorderBrush = Brushes.Gray };
    private readonly TextBlock _error = new() { Name = "ColorError", Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap };
    private readonly Button _ok = new() { Name = "ColorConfirm", Content = UiText.Get("确定"), IsDefault = true, MinWidth = 80 };
    private readonly Slider[] _sliders = new Slider[4];
    private readonly NumericUpDown[] _numbers = new NumericUpDown[4];
    private readonly ColorSpectrum _spectrum = new() { Name = "ColorSpectrum", Height = 240 };
    private bool _syncing;

    public CustomColorDialog(string? initialColor, string? title = null)
    {
        Title = title ?? UiText.Get("自定义颜色"); Width = 720; SizeToContent = SizeToContent.Height; CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var panel = new StackPanel { Margin = new(20), Spacing = 12 };
        var editor = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,20,280") };
        var picking = new StackPanel { Spacing = 8 };
        picking.Children.Add(new TextBlock { Text = UiText.Get("色盘 · 拖动选取颜色") }); picking.Children.Add(_spectrum);
        picking.Children.Add(new TextBlock { Text = UiText.Get("下方色相条切换色系"), Foreground = Brushes.Gray });
        var inputs = new StackPanel { Spacing = 8 };
        global::Avalonia.Controls.Grid.SetColumn(inputs, 2); editor.Children.Add(picking); editor.Children.Add(inputs); panel.Children.Add(editor);
        var previews = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,12,*"), RowDefinitions = new("Auto,8,52") };
        previews.Children.Add(new TextBlock { Text = UiText.Get("原颜色") });
        var label = new TextBlock { Text = UiText.Get("新颜色") }; global::Avalonia.Controls.Grid.SetColumn(label, 2); previews.Children.Add(label);
        var original = ColorPalette.Swatch(initialColor); original.Width = double.NaN; original.Height = double.NaN;
        if (initialColor is null) original.Child = new TextBlock { Text = UiText.Get("默认颜色"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        global::Avalonia.Controls.Grid.SetRow(original, 2); previews.Children.Add(original);
        global::Avalonia.Controls.Grid.SetRow(_preview, 2); global::Avalonia.Controls.Grid.SetColumn(_preview, 2); previews.Children.Add(_preview);
        AutomationProperties.SetName(_preview, UiText.Get("新颜色预览")); picking.Children.Add(previews);
        panel.Children.Add(new TextBlock { Text = UiText.Get("常用颜色") });
        var palette = new WrapPanel { Orientation = Orientation.Horizontal };
        foreach (var (name, value) in ColorPalette.Common)
        {
            var button = new Button { Content = ColorPalette.Swatch(value, 24), Padding = new(5), Margin = new(0, 0, 4, 4) };
            ToolTip.SetTip(button, name); AutomationProperties.SetName(button, name);
            button.Click += (_, _) => SetColor(Color.Parse("#" + value)); palette.Children.Add(button);
        }
        panel.Children.Add(palette);
        inputs.Children.Add(new TextBlock { Text = UiText.Get("十六进制色值") });
        AutomationProperties.SetName(_hex, UiText.Get("十六进制颜色")); inputs.Children.Add(_hex);
        inputs.Children.Add(new TextBlock { Text = UiText.Get("#RRGGBB 或 AARRGGBB"), Foreground = Brushes.Gray });
        string[] names = [UiText.Get("红 R"), UiText.Get("绿 G"), UiText.Get("蓝 B"), UiText.Get("不透明度 A")];
        for (var i = 0; i < names.Length; i++)
        {
            var index = i;
            var row = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,76"), RowDefinitions = new("32,Auto") };
            row.Children.Add(new TextBlock { Text = names[i], VerticalAlignment = VerticalAlignment.Center });
            var slider = _sliders[i] = new Slider { Minimum = 0, Maximum = 255, TickFrequency = 1, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
            var number = _numbers[i] = new NumericUpDown { Name = "ColorChannel" + i, Minimum = 0, Maximum = 255, Increment = 1, FormatString = "0", Height = 32, MinHeight = 32, ShowButtonSpinner = false };
            AutomationProperties.SetName(slider, names[i]); AutomationProperties.SetName(number, names[i] + UiText.Get(" 数值"));
            global::Avalonia.Controls.Grid.SetRow(slider, 1); global::Avalonia.Controls.Grid.SetColumnSpan(slider, 2); global::Avalonia.Controls.Grid.SetColumn(number, 1);
            row.Children.Add(slider); row.Children.Add(number); inputs.Children.Add(row);
            slider.PropertyChanged += (_, e) => { if (e.Property == Slider.ValueProperty && !_syncing) SetChannel(index, (byte)Math.Round(slider.Value)); };
            number.ValueChanged += (_, _) => { if (!_syncing && number.Value is { } value) SetChannel(index, (byte)Math.Clamp(decimal.Round(value), 0, 255)); };
        }
        _hex.TextChanged += (_, _) =>
        {
            if (_syncing) return;
            if (ColorPalette.TryParse(_hex.Text, out var color)) SetColor(color, updateText: false);
            else { _error.Text = UiText.Get("请输入 6 位 RGB 或 8 位 ARGB 十六进制色值。"); _ok.IsEnabled = false; }
        };
        _spectrum.ColorChanged += color => { if (!_syncing) SetColor(color); };
        panel.Children.Add(_error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        var cancel = new Button { Name = "ColorCancel", Content = UiText.Get("取消"), IsCancel = true, MinWidth = 80 };
        _ok.Click += (_, _) => { if (ColorPalette.TryParse(_hex.Text, out var color)) Close(color.ToString()[1..].ToUpperInvariant()); };
        cancel.Click += (_, _) => Close(null);
        buttons.Children.Add(_ok); buttons.Children.Add(cancel); panel.Children.Add(buttons); Content = panel;
        SetColor(ColorPalette.TryParse(initialColor, out var initial) ? initial : Colors.Black);
        Opened += (_, _) => { _hex.Focus(); _hex.SelectAll(); };
    }

    private void SetChannel(int index, byte value)
    {
        var channels = _sliders.Select(s => (byte)Math.Round(s.Value)).ToArray(); channels[index] = value;
        SetColor(Color.FromArgb(channels[3], channels[0], channels[1], channels[2]));
    }

    private void SetColor(Color color, bool updateText = true)
    {
        _syncing = true;
        try
        {
            if (updateText) _hex.Text = color.ToString().ToUpperInvariant();
            _spectrum.SelectedColor = color;
            byte[] values = [color.R, color.G, color.B, color.A];
            for (var i = 0; i < values.Length; i++) { _sliders[i].Value = values[i]; _numbers[i].Value = values[i]; }
            _preview.Background = new SolidColorBrush(color); _error.Text = ""; _ok.IsEnabled = true;
        }
        finally { _syncing = false; }
    }

    public Task<string?> ShowAsync(Window owner) => owner is MainWindow main ? main.ShowOwnedAsync<string?>(this) : ShowDialog<string?>(owner);
}
