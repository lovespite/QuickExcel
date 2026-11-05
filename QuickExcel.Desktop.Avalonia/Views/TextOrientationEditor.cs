using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class TextOrientationEditor : UserControl
{
    private readonly NumericUpDown _degrees = new()
    {
        Name = "RotationDegrees", Minimum = -90, Maximum = 90, Increment = 1,
        FormatString = "0", Width = 120, HorizontalAlignment = HorizontalAlignment.Left
    };
    private readonly CheckBox _stacked = new() { Name = "RotationStacked", Content = UiText.Get("竖排文字") };
    private readonly TextBlock _preview = new()
    {
        Name = "RotationPreview", FontSize = 16, HorizontalAlignment = HorizontalAlignment.Center,
        VerticalAlignment = VerticalAlignment.Center, TextAlignment = TextAlignment.Center,
        RenderTransformOrigin = RelativePoint.Center
    };

    public TextOrientationEditor(int rotation)
    {
        var body = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,180") };
        var options = new StackPanel { Spacing = 6 };
        options.Children.Add(new TextBlock { Text = UiText.Get("文字方向") });
        options.Children.Add(_stacked);
        options.Children.Add(new TextBlock { Text = UiText.Get("角度（−90° 到 90°）") });
        options.Children.Add(_degrees); body.Children.Add(options);
        var preview = new Border { Height = 140, BorderThickness = new(1), BorderBrush = Brushes.LightGray, ClipToBounds = true, Child = _preview };
        global::Avalonia.Controls.Grid.SetColumn(preview, 1); body.Children.Add(preview); Content = body;
        AutomationProperties.SetName(_degrees, UiText.Get("文字旋转角度"));
        _degrees.Value = TextRotationOptions.ToDegrees(rotation); _stacked.IsChecked = rotation == 255;
        _degrees.ValueChanged += (_, _) => Refresh(); _stacked.IsCheckedChanged += (_, _) => Refresh(); Refresh();
    }

    public int Rotation => _stacked.IsChecked == true ? 255 : _degrees.Value is { } value && value is >= -90 and <= 90 && value == decimal.Truncate(value)
        ? TextRotationOptions.FromDegrees((int)value) : throw new ArgumentException(UiText.Get("请输入 −90 到 90 之间的整数角度。"));

    private void Refresh()
    {
        var stacked = _stacked.IsChecked == true; _degrees.IsEnabled = !stacked;
        var text = UiText.Get("文字示例");
        _preview.Text = stacked ? string.Join("\n", System.Globalization.StringInfo.ParseCombiningCharacters(text).Select(start =>
            System.Globalization.StringInfo.GetNextTextElement(text, start))) : text;
        _preview.RenderTransform = new RotateTransform(stacked ? 0 : -(double)(_degrees.Value ?? 0));
    }
}
