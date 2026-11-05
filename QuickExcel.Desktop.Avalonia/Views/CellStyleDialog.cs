using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class CellStyleDialog : Window
{
    private readonly TextBox _name = new() { Name = "CellStyleName", MaxLength = 64 };
    private readonly TextBlock _error = new() { Name = "CellStyleError", Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap };
    private readonly Dictionary<CellStyleParts, CheckBox> _parts = [];
    private readonly Border _preview = new() { Padding = new(12), MinHeight = 55, BorderThickness = new(1), BorderBrush = Brushes.LightGray };
    private StyleDefinition _format;
    private readonly QuickCellStyle _original;

    internal CellStyleDialog(QuickCellStyle style, CellValue sample, ExcelDateSystem dateSystem, IEnumerable<StyleDefinition> styles, Action<QuickCellStyle> validate)
    {
        _original = style; _format = style.Format;
        Title = UiText.Get("样式"); Width = 480; Height = 650; MinWidth = 430; MinHeight = 590;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var root = new global::Avalonia.Controls.Grid { RowDefinitions = new("*,Auto,Auto"), Margin = new(22) };
        var panel = new StackPanel { Spacing = 12 };
        panel.Children.Add(new TextBlock { Text = UiText.Get("样式名") }); _name.Text = style.Name; panel.Children.Add(_name);
        AutomationProperties.SetName(_name, UiText.Get("样式名"));
        var format = new Button { Name = "CellStyleFormat", Content = UiText.Get("格式…"), HorizontalAlignment = HorizontalAlignment.Right, MinWidth = 90 };
        panel.Children.Add(format); panel.Children.Add(_preview);
        panel.Children.Add(new TextBlock { Text = UiText.Get("样式包括（举例）"), FontWeight = FontWeight.SemiBold });
        foreach (var part in new[] { CellStyleParts.Number, CellStyleParts.Alignment, CellStyleParts.Font, CellStyleParts.Border, CellStyleParts.Fill, CellStyleParts.Protection })
        {
            var check = new CheckBox { Name = "CellStylePart" + part, IsChecked = style.Parts.HasFlag(part), HorizontalAlignment = HorizontalAlignment.Stretch };
            _parts.Add(part, check); panel.Children.Add(check);
        }
        panel.Children.Add(new TextBlock { Text = UiText.Get("保存在本机，所有工作簿均可使用。取消勾选的类别不会改变目标单元格。"), TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Gray, FontSize = 12 });
        root.Children.Add(new ScrollViewer { Content = panel });
        global::Avalonia.Controls.Grid.SetRow(_error, 1); _error.Margin = new(0, 8); root.Children.Add(_error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        var ok = new Button { Name = "CellStyleConfirm", Content = UiText.Get("确定"), IsDefault = true, MinWidth = 90 };
        var cancel = new Button { Name = "CellStyleCancel", Content = UiText.Get("取消"), IsCancel = true, MinWidth = 90 };
        buttons.Children.Add(ok); buttons.Children.Add(cancel); global::Avalonia.Controls.Grid.SetRow(buttons, 2); root.Children.Add(buttons); Content = root;
        cancel.Click += (_, _) => Close(null);
        ok.Click += (_, _) =>
        {
            try
            {
                var result = _original with { Name = (_name.Text ?? "").Trim(), Format = _format, Parts = _parts.Where(p => p.Value.IsChecked == true).Aggregate(CellStyleParts.None, (value, pair) => value | pair.Key) };
                validate(result); Close(result);
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { _error.Text = ex.Message; }
        };
        format.Click += async (_, _) =>
        {
            format.IsEnabled = false;
            try
            {
                var s = _format;
                BorderStroke?[] borders = [new(s.LeftBorder, s.LeftBorderColor), new(s.RightBorder, s.RightBorderColor), new(s.TopBorder, s.TopBorderColor), new(s.BottomBorder, s.BottomBorderColor), new(0), new(0), new(s.DiagonalUp ? s.DiagonalBorder : 0, s.DiagonalBorderColor), new(s.DiagonalDown ? s.DiagonalBorder : 0, s.DiagonalBorderColor)];
                var dialog = new NumberFormatDialog(s, sample, dateSystem, styles, true, borders);
                if (await dialog.ShowAsync(this) is { } patch) { _format = QuickCellStyle.EditFormat(s, patch, dialog.BorderPatch); Refresh(); }
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { _error.Text = ex.Message; }
            finally { format.IsEnabled = true; }
        };
        Refresh(); Opened += (_, _) => { _name.Focus(); _name.SelectAll(); };
    }

    private void Refresh()
    {
        var s = _format;
        void Label(CellStyleParts part, string text) { _parts[part].Content = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap }; AutomationProperties.SetName(_parts[part], text); }
        Label(CellStyleParts.Number, UiText.Get("数字") + ": " + NumberFormatOptions.Label(s));
        Label(CellStyleParts.Alignment, UiText.Get("对齐") + ": " + new[] { UiText.Get("常规"), UiText.Get("左对齐"), UiText.Get("居中"), UiText.Get("右对齐"), UiText.Get("填充") }[s.Horizontal] + ", " + new[] { UiText.Get("顶端对齐"), UiText.Get("垂直居中"), UiText.Get("底端对齐"), UiText.Get("两端对齐") }[s.Vertical]);
        Label(CellStyleParts.Font, UiText.Get("字体") + $": {s.FontName ?? "Microsoft YaHei UI"} {s.FontSize ?? 11}" + (s.Bold ? " · " + UiText.Get("加粗") : "") + (s.Italic ? " · " + UiText.Get("倾斜") : ""));
        Label(CellStyleParts.Border, UiText.Get("边框") + ": " + (s.LeftBorder + s.RightBorder + s.TopBorder + s.BottomBorder + s.DiagonalBorder == 0 ? UiText.Get("无框线") : UiText.Get("自定义边框")));
        Label(CellStyleParts.Fill, UiText.Get("填充") + ": " + (s.FillColor is null ? UiText.Get("无填充") : "#" + s.FillColor[2..]));
        Label(CellStyleParts.Protection, UiText.Get("保护") + ": " + (s.Locked ? UiText.Get("锁定") : UiText.Get("未锁定")) + (s.FormulaHidden ? " · " + UiText.Get("隐藏公式") : ""));
        _preview.Child = MainWindow.CellStyleSample(_original with { Format = s }, UiText.Get("样式预览"));
    }
}
