using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class BorderEditor : UserControl
{
    internal static readonly string[] StyleNames = [UiText.Get("无"), UiText.Get("细实线"), UiText.Get("中实线"), UiText.Get("虚线"), UiText.Get("点线"), UiText.Get("粗实线"), UiText.Get("双线"), UiText.Get("发丝线"), UiText.Get("中虚线"), UiText.Get("点划线"), UiText.Get("中点划线"), UiText.Get("双点划线"), UiText.Get("中双点划线"), UiText.Get("斜点划线")];
    internal static readonly (BorderPreset Preset, string Name)[] Presets = [
        (BorderPreset.Bottom, UiText.Get("下框线")), (BorderPreset.Top, UiText.Get("上框线")), (BorderPreset.Left, UiText.Get("左框线")), (BorderPreset.Right, UiText.Get("右框线")),
        (BorderPreset.None, UiText.Get("无框线")), (BorderPreset.All, UiText.Get("所有框线")), (BorderPreset.Outside, UiText.Get("外侧框线")), (BorderPreset.ThickOutside, UiText.Get("粗外侧框线")),
        (BorderPreset.DoubleBottom, UiText.Get("双底框线")), (BorderPreset.ThickBottom, UiText.Get("粗下框线")), (BorderPreset.TopBottom, UiText.Get("上下框线")),
        (BorderPreset.TopThickBottom, UiText.Get("上框线和粗下框线")), (BorderPreset.TopDoubleBottom, UiText.Get("上框线和双下框线")),
        (BorderPreset.Inside, UiText.Get("内部框线")), (BorderPreset.InsideHorizontal, UiText.Get("内部横框线")), (BorderPreset.InsideVertical, UiText.Get("内部竖框线"))];
    private readonly BorderPreview _preview = new() { Name = "BorderPreview", Height = 180, SampleText = true };
    private readonly ListBox _styles = new() { Name = "BorderStyles", Height = 350 };
    private string? _color;
    private readonly bool[] _changed = new bool[8];
    private readonly Button[] _edges = new Button[8];
    private readonly BorderStroke?[] _initial;
    public CellBorderPatch? Patch
    {
        get
        {
            if (!_changed.Any(c => c)) return null;
            BorderStroke? Edge(int i) => _changed[i] ? _preview.Lines[i] : null;
            var diagonal = _changed[6] && _preview.Lines[6]?.Style > 0 ? _preview.Lines[6] : _changed[7] && _preview.Lines[7]?.Style > 0 ? _preview.Lines[7] : null;
            return new(Edge(0), Edge(1), Edge(2), Edge(3), Edge(4), Edge(5), diagonal,
                _changed[6] ? _preview.Lines[6]?.Style > 0 : null, _changed[7] ? _preview.Lines[7]?.Style > 0 : null);
        }
    }
    public BorderEditor(BorderStroke?[] initial)
    {
        _initial = initial; _preview.Lines = initial.ToArray();
        var root = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("180,20,*"), Margin = new(4, 16) };
        var choices = new StackPanel { Spacing = 8 }; root.Children.Add(choices);
        choices.Children.Add(new TextBlock { Text = UiText.Get("线型"), FontWeight = FontWeight.SemiBold });
        _styles.ItemsSource = Enumerable.Range(0, StyleNames.Length).ToArray(); _styles.SelectedIndex = 1;
        _styles.ItemsPanel = new FuncTemplate<Panel?>(() => new UniformGrid { Columns = 2 });
        _styles.ItemTemplate = new FuncDataTemplate<int>((kind, _) =>
        {
            var row = new StackPanel { Spacing = 2 };
            row.Children.Add(new BorderLineSample(kind) { Width = 60, Height = 12 }); row.Children.Add(new TextBlock { Text = StyleNames[kind], FontSize = 11, HorizontalAlignment = HorizontalAlignment.Center }); return row;
        });
        AutomationProperties.SetName(_styles, UiText.Get("边框线型")); choices.Children.Add(_styles);
        choices.Children.Add(new TextBlock { Text = UiText.Get("颜色"), FontWeight = FontWeight.SemiBold });
        var colorMenu = new MenuFlyout(); var colorButton = new DropDownButton { Name = "BorderColor", Content = UiText.Get("自动（黑色）"), Flyout = colorMenu, HorizontalAlignment = HorizontalAlignment.Stretch };
        void Color(string? value, string name) { _color = value; colorButton.Content = name; }
        var colors = new List<Control>(); var automatic = new MenuItem { Header = UiText.Get("自动（黑色）") }; automatic.Click += (_, _) => Color(null, UiText.Get("自动（黑色）")); colors.Add(automatic);
        foreach (var (name, value) in ColorPalette.Common) { var item = new MenuItem { Header = name, Icon = ColorPalette.Swatch(value) }; item.Click += (_, _) => Color(value, name); colors.Add(item); }
        var custom = new MenuItem { Header = UiText.Get("自定义颜色…") }; custom.Click += async (_, _) =>
        {
            if (TopLevel.GetTopLevel(this) is Window owner && await new CustomColorDialog(_color, UiText.Get("边框颜色")).ShowDialog<string?>(owner) is { } color) Color(color, "#" + color[2..]);
        };
        colors.Add(custom); colorMenu.ItemsSource = colors; choices.Children.Add(colorButton);
        var detail = new StackPanel { Spacing = 8 }; global::Avalonia.Controls.Grid.SetColumn(detail, 2); root.Children.Add(detail);
        detail.Children.Add(new TextBlock { Text = UiText.Get("预置"), FontWeight = FontWeight.SemiBold });
        var presets = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        foreach (var (preset, label) in new[] { (BorderPreset.None, UiText.Get("无")), (BorderPreset.Outside, UiText.Get("外边框")), (BorderPreset.Inside, UiText.Get("内部")) })
        {
            var button = new Button { Name = "BorderPreset" + preset, Content = label, MinWidth = 90 }; button.Click += (_, _) => Apply(CellBorders.Preset(preset, _styles.SelectedIndex, _color)); presets.Children.Add(button);
        }
        detail.Children.Add(presets); detail.Children.Add(new TextBlock { Text = UiText.Get("边框预览 · 点击线段设置或取消") });
        _preview.Toggle = Toggle; detail.Children.Add(_preview);
        var buttons = new WrapPanel { Orientation = Orientation.Horizontal };
        var labels = new[] { UiText.Get("左"), UiText.Get("右"), UiText.Get("上"), UiText.Get("下"), UiText.Get("内部横线"), UiText.Get("内部竖线"), UiText.Get("斜上框线 ↗"), UiText.Get("斜下框线 ↘") };
        for (var i = 0; i < 8; i++)
        {
            var edge = i; var button = new Button { Name = "BorderEdge" + i, Content = labels[i], Margin = new(0, 0, 6, 6), Padding = new(8, 4) };
            AutomationProperties.SetName(button, labels[i]); button.Click += (_, _) => Toggle(edge); _edges[i] = button; buttons.Children.Add(button);
        }
        detail.Children.Add(buttons); detail.Children.Add(new TextBlock { Text = UiText.Get("先选择线型和颜色，再点击预置或预览中的边框。灰色虚线表示混合状态；未操作的边框保持原样。"), TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Gray, FontSize = 12 });
        var reset = new Button { Name = "BorderReset", Content = UiText.Get("恢复原有边框") }; reset.Click += (_, _) => { _preview.Lines = _initial.ToArray(); Array.Clear(_changed); Refresh(); }; detail.Children.Add(reset);
        Content = root; Refresh();
    }
    private void Toggle(int edge)
    {
        var next = new BorderStroke(Math.Max(0, _styles.SelectedIndex), _color);
        _preview.Lines[edge] = _preview.Lines[edge] is { } old && old.Style == next.Style &&
            string.Equals(old.Color ?? "FF000000", next.Color ?? "FF000000", StringComparison.OrdinalIgnoreCase) ? new(0) : next; _changed[edge] = true;
        if (edge >= 6 && next.Style > 0 && _preview.Lines[edge]?.Style > 0 && _preview.Lines[13 - edge]?.Style > 0)
        { _preview.Lines[13 - edge] = next; _changed[13 - edge] = true; }
        Refresh();
    }
    private void Apply(CellBorderPatch patch)
    {
        var values = new[] { patch.Left, patch.Right, patch.Top, patch.Bottom, patch.InsideHorizontal, patch.InsideVertical, patch.Diagonal, patch.Diagonal };
        for (var i = 0; i < 8; i++) if (values[i] is { } stroke) { _preview.Lines[i] = stroke; _changed[i] = true; }
        Refresh();
    }
    private void Refresh()
    {
        _preview.InvalidateVisual();
        for (var i = 0; i < _edges.Length; i++) if (_edges[i] is { } button) button.Background = _preview.Lines[i]?.Style > 0 ? Brush.Parse("#DCEFE8") : Brushes.Transparent;
    }
}
