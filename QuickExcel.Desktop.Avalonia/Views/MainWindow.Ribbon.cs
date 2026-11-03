using System.Globalization;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly TabControl _ribbon = new() { Height = 170 };
    // Fluent's default 32-DIP minimum overflows the 30-DIP ribbon row and clips its top border.
    private readonly ComboBox _font = new() { Width = 160, Height = 28, MinHeight = 28, Padding = new(8, 2), VerticalContentAlignment = VerticalAlignment.Center },
        _fontSize = new() { Width = 72, Height = 28, MinHeight = 28, Padding = new(8, 2), VerticalContentAlignment = VerticalAlignment.Center };
    private readonly List<Control> _mutationControls = [];
    private readonly Dictionary<string, StackPanel> _ribbonPages = [];
    private bool _updatingRibbon;
    private WorkbookClipboard? _formatBrush;
    private Guid _formatBrushSession;
    private StyleDefinition CurrentStyle => Document.Snapshot.Styles[Document.Snapshot.Sheets[SheetIndex].GetStyleId(_grid.Current.Row, _grid.Current.Column)];
    private Button RibbonButton(string title, Func<Task> action, bool mutation = true)
    {
        var button = new Button { Content = title, Padding = new(8, 3), Margin = new(2, 1), Height = 28,
            Background = Brushes.Transparent, BorderThickness = new(0),
            HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Left };
        AutomationProperties.SetName(button, title);
        button.Click += async (_, _) => await PerformAsync(action);
        if (mutation) _mutationControls.Add(button); return button;
    }
    private void RibbonGroup(string page, string name, params Control[] controls)
    {
        var body = new global::Avalonia.Controls.Grid { RowDefinitions = new("30,30,30") };
        for (var i = 0; i < (controls.Length + 2) / 3; i++) body.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        for (var i = 0; i < controls.Length; i++)
        {
            var control = controls[i]; control.VerticalAlignment = VerticalAlignment.Center;
            global::Avalonia.Controls.Grid.SetRow(control, i % 3); global::Avalonia.Controls.Grid.SetColumn(control, i / 3); body.Children.Add(control);
        }
        if (name == "字体")
        {
            body.Children.Clear(); body.Width = 244; body.ColumnDefinitions = new("*,*");
            var selectors = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            selectors.Children.Add(_font); selectors.Children.Add(_fontSize);
            global::Avalonia.Controls.Grid.SetColumnSpan(selectors, 2); body.Children.Add(selectors);
            for (var i = 2; i < controls.Length; i++)
            {
                global::Avalonia.Controls.Grid.SetRow(controls[i], i / 2); global::Avalonia.Controls.Grid.SetColumn(controls[i], i % 2); body.Children.Add(controls[i]);
            }
        }
        RibbonContentGroup(page, name, body);
    }
    private void RibbonContentGroup(string page, string name, Control body)
    {
        if (!_ribbonPages.TryGetValue(page, out var row))
        {
            row = new StackPanel { Orientation = Orientation.Horizontal };
            _ribbonPages.Add(page, row);
            var scroll = new ScrollViewer { Content = row, VerticalAlignment = VerticalAlignment.Stretch, HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled };
            scroll.PointerWheelChanged += (_, e) =>
            {
                if (scroll.Extent.Width <= scroll.Viewport.Width || e.Delta.Y == 0) return;
                scroll.Offset = new(Math.Clamp(scroll.Offset.X - e.Delta.Y * 72, 0, scroll.Extent.Width - scroll.Viewport.Width), 0); e.Handled = true;
            };
            _ribbon.Items.Add(new TabItem { Header = page, FontSize = 14, Padding = new(12, 6), Content = scroll });
        }
        var panel = new global::Avalonia.Controls.Grid { RowDefinitions = new("90,20"), Margin = new(7, 2) };
        var label = new TextBlock { Text = name, FontSize = 11, Foreground = Brushes.Gray, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        global::Avalonia.Controls.Grid.SetRow(label, 1); panel.Children.Add(body); panel.Children.Add(label);
        row.Children.Add(new Border { BorderBrush = Brushes.LightGray, BorderThickness = new(0, 0, 1, 0), Child = panel });
    }
    private void ArrangeRibbonTabs()
    {
        string[] order = ["开始", "插入", "页面布局", "公式", "自动执行", "数据", "审阅", "视图"];
        var tabs = _ribbon.Items.Cast<TabItem>().OrderBy(t => Array.IndexOf(order, t.Header as string)).ToArray();
        _ribbon.Items.Clear(); foreach (var tab in tabs) _ribbon.Items.Add(tab); _ribbon.SelectedIndex = 0;
    }
    private void InstallRibbon()
    {
        AutomationProperties.SetName(_ribbon, "功能区"); DockPanel.SetDock(_ribbon, Dock.Top); _book.Children.Insert(0, _ribbon);
        RibbonGroup("开始", "剪贴板", RibbonButton("粘贴", () => PasteAsync("all")), RibbonButton("复制", () => CopyAsync(false), false),
            RibbonButton("剪切", () => CopyAsync(true)), RibbonButton("只贴值", () => PasteAsync("values")), RibbonButton("格式刷", FormatBrushAsync));
        _font.ItemsSource = FontManager.Current.SystemFonts.Select(f => f.Name).Order(StringComparer.CurrentCultureIgnoreCase).ToArray();
        _fontSize.ItemsSource = new double[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72 };
        AutomationProperties.SetName(_font, "字体"); AutomationProperties.SetName(_fontSize, "字号");
        _font.SelectionChanged += async (_, _) => { if (!_updatingRibbon && _font.SelectedItem is string name) await PerformAsync(() => ApplyStyleAsync(new(FontName: name))); };
        _fontSize.SelectionChanged += async (_, _) => { if (!_updatingRibbon && _fontSize.SelectedItem is double size) await PerformAsync(() => ApplyStyleAsync(new(FontSize: size))); };
        _mutationControls.AddRange([_font, _fontSize]);
        RibbonGroup("开始", "字体", _font, _fontSize, RibbonButton("加粗", () => ApplyStyleAsync(new(Bold: !CurrentStyle.Bold))),
            RibbonButton("倾斜", () => ApplyStyleAsync(new(Italic: !CurrentStyle.Italic))), RibbonButton("字体颜色", () => ColorAsync(false)), RibbonButton("填充颜色", () => ColorAsync(true)));
        RibbonGroup("开始", "对齐", RibbonButton("左对齐", () => ApplyStyleAsync(new(Horizontal: 1))), RibbonButton("居中", () => ApplyStyleAsync(new(Horizontal: 2))),
            RibbonButton("右对齐", () => ApplyStyleAsync(new(Horizontal: 3))), RibbonButton("自动换行", () => ApplyStyleAsync(new(WrapText: !CurrentStyle.WrapText))),
            RibbonButton("合并", () => MergeAsync(false)), RibbonButton("取消合并", () => MergeAsync(true)));
        RibbonGroup("开始", "数字与单元格", RibbonButton("数字格式", NumberFormatAsync), RibbonButton("增加小数位", () => PrecisionAsync(1)),
            RibbonButton("减少小数位", () => PrecisionAsync(-1)), RibbonButton("设置单元格格式", CellFormatAsync), RibbonButton("清除格式", ClearFormatsAsync));
        RibbonGroup("公式", "公式", RibbonButton("自动求和  Alt+=", () => { BeginAutoSum(); return Task.CompletedTask; }), RibbonButton("重算  F9", CalculateAsync, false));
        RibbonGroup("开始", "编辑", RibbonButton("向下填充", () => FillDirectionAsync(false)), RibbonButton("向右填充", () => FillDirectionAsync(true)), RibbonButton("序列", SeriesAsync),
            RibbonButton("转为普通文本", FlattenRichTextAsync));
        _grid.AutoFitRequested += (axis, index) => _ = PerformAsync(() => AutoFitAsync(axis, index));
        RibbonGroup("视图", "行列尺寸", RibbonButton("自动行高", () => AutoFitAsync(SheetAxis.Row, _grid.Current.Row)),
            RibbonButton("自动列宽", () => AutoFitAsync(SheetAxis.Column, _grid.Current.Column)), RibbonButton("行高", () => AxisSizeAsync(SheetAxis.Row)), RibbonButton("列宽", () => AxisSizeAsync(SheetAxis.Column)));
    }
    private void UpdateRibbon()
    {
        _updatingRibbon = true;
        try
        {
            _ribbon.IsEnabled = !NonExecutionBusy;
            foreach (var control in _mutationControls) control.IsEnabled = Document.Snapshot.FileState?.FormulaReadOnly != true;
            var style = CurrentStyle; var name = style.FontName ?? "Microsoft YaHei UI"; var size = style.FontSize ?? 11;
            if (!_font.Items.Cast<string>().Contains(name)) _font.ItemsSource = _font.Items.Cast<string>().Append(name).Order().ToArray();
            if (!_fontSize.Items.Cast<double>().Contains(size)) _fontSize.ItemsSource = _fontSize.Items.Cast<double>().Append(size).Order().ToArray();
            _font.SelectedItem = name; _fontSize.SelectedItem = size;
        }
        finally { _updatingRibbon = false; }
        RefreshScriptActions();
    }
    private CellAddress[] VisibleAddresses()
    {
        var snapshot = Document.Snapshot; var view = CurrentView; view.RequireCurrentFilter(snapshot);
        return SelectedAddresses().Where(p => !view.IsFilteredOut(snapshot, SheetIndex, p.Row)).ToArray();
    }
    private async Task ApplyStyleAsync(CellStylePatch patch)
    { if (await CommitEditorsAsync()) await ExecuteAsync(new SetCellStyleCommand(VisibleAddresses(), patch)); }
    private async Task PrecisionAsync(int delta)
    { if (await CommitEditorsAsync()) await ExecuteAsync(new AdjustDecimalPlacesCommand(VisibleAddresses(), delta)); }
    private async Task ClearFormatsAsync()
    { if (await CommitEditorsAsync()) await ExecuteAsync(new ClearCellFormatsCommand(VisibleAddresses())); }
    private async Task MergeAsync(bool remove)
    { if (await CommitEditorsAsync()) await ExecuteAsync(new MergeCellsCommand(Document.Snapshot.ActiveSheetId, Rectangle(SelectedRange()), remove)); }
    private async Task ColorAsync(bool fill)
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses();
        var dialog = new FieldsDialog(fill ? "填充颜色" : "字体颜色").Choice("kind", "颜色", ["自定义", "自动 / 无填充", "黑色", "白色", "红色", "绿色", "蓝色", "黄色"], 0)
            .Text("color", "自定义色值（#RRGGBB 或 AARRGGBB）", (fill ? CurrentStyle.FillColor : CurrentStyle.FontColor) ?? "FF000000");
        string Value()
        {
            var value = dialog.Index("kind") switch { 2 => "FF000000", 3 => "FFFFFFFF", 4 => "FFFF0000", 5 => "FF008000", 6 => "FF0000FF", 7 => "FFFFFF00", _ => dialog.String("color").Trim().TrimStart('#') };
            if (value.Length == 6) value = "FF" + value;
            if (dialog.Index("kind") != 1 && (value.Length != 8 || value.Any(c => !Uri.IsHexDigit(c)))) throw new ArgumentException("请输入有效的十六进制颜色。");
            return value.ToUpperInvariant();
        }
        dialog.Validate = () => Value(); if (!await dialog.ShowAsync(this)) return;
        var clear = dialog.Index("kind") == 1;
        await ExecuteAsync(new SetCellStyleCommand(addresses, fill ? new(FillColor: clear ? null : Value(), ClearFillColor: clear) : new(FontColor: clear ? null : Value(), ClearFontColor: clear)), revision);
    }
    private async Task NumberFormatAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses();
        var dialog = new FieldsDialog("数字格式").Choice("kind", "分类", ["常规", "数字", "千位分隔", "百分比", "科学计数", "日期", "货币", "自定义"])
            .Text("code", "自定义格式代码", CurrentStyle.NumberFormatCode ?? "#,##0.00");
        if (!await dialog.ShowAsync(this)) return;
        var patch = dialog.Index("kind") switch
        {
            1 => new CellStylePatch(NumberFormatId: 2), 2 => new(NumberFormatId: 4), 3 => new(NumberFormatId: 10), 4 => new(NumberFormatId: 11),
            5 => new(NumberFormatId: 0, NumberFormatCode: "yyyy-mm-dd"), 6 => new(NumberFormatId: 0, NumberFormatCode: "\"¥\"#,##0.00;[Red](\"¥\"#,##0.00)"),
            7 => new(NumberFormatId: 0, NumberFormatCode: dialog.String("code")), _ => new(NumberFormatId: 0)
        };
        await ExecuteAsync(new SetCellStyleCommand(addresses, patch), revision);
    }
    private async Task CellFormatAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses(); var style = CurrentStyle;
        var dialog = new FieldsDialog("单元格格式").Choice("horizontal", "水平对齐", ["常规", "左对齐", "居中", "右对齐"], Math.Clamp(style.Horizontal, 0, 3))
            .Choice("vertical", "垂直对齐", ["顶端", "居中", "底端", "两端对齐"], style.Vertical)
            .Choice("border", "边框", ["保持原样", "四周细边框", "无边框"])
            .Check("wrap", "自动换行", style.WrapText).Check("bold", "加粗", style.Bold).Check("italic", "倾斜", style.Italic)
            .Text("size", "字号", (style.FontSize ?? 11).ToString(CultureInfo.CurrentCulture));
        dialog.Validate = () => dialog.Number("size", .1, 409); if (!await dialog.ShowAsync(this)) return;
        await ExecuteAsync(new SetCellStyleCommand(addresses, new(Horizontal: dialog.Index("horizontal"), Vertical: dialog.Index("vertical"),
            Border: dialog.Index("border") == 0 ? null : dialog.Index("border") == 1, WrapText: dialog.Boolean("wrap"), Bold: dialog.Boolean("bold"), Italic: dialog.Boolean("italic"), FontSize: dialog.Number("size", .1, 409))), revision);
    }
    private async Task FormatBrushAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot;
        if (_formatBrush is null)
        {
            _formatBrush = WorkbookClipboard.Capture(snapshot, snapshot.ActiveSheetId, new(_grid.Current.Row, _grid.Current.Column, _grid.Current.Row, _grid.Current.Column));
            _formatBrushSession = Document.Session.SessionId; _status.Text = "已采样格式，选择目标后再次点击格式刷；Esc 取消。"; return;
        }
        var brush = _formatBrush; _formatBrush = null;
        if (_formatBrushSession != Document.Session.SessionId) throw new InvalidOperationException("原工作簿已关闭，请重新采样格式。");
        var range = SelectedRange(); RequireVisiblePaste(range);
        await ExecuteAsync(new PasteCellsCommand(brush, snapshot.ActiveSheetId, Rectangle(range), brush.Revision, CellPasteMode.Formats), brush.Revision);
    }
    private async Task FlattenRichTextAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses();
        if (await Dialogs.Confirm(this, "转换为普通文本", "将移除选中单元格的分段字体格式，保留文字内容。可撤销。"))
            await ExecuteAsync(new FlattenRichTextCommand(addresses), revision);
    }
    private async Task AutoFitAsync(SheetAxis axis, int index)
    {
        if (!await CommitEditorsAsync()) return;
        var pixels = _grid.MeasureAutoFit(axis, index);
        await ExecuteAsync(new SetAxisMetadataCommand(Document.Snapshot.ActiveSheetId, axis, index, 1, axis == SheetAxis.Row ? pixels : (pixels - 5) / 7));
    }
    private async Task AxisSizeAsync(SheetAxis axis)
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var range = SelectedRange(); var sheet = Document.Snapshot.Sheets[SheetIndex];
        var row = axis == SheetAxis.Row; var size = row ? DrawingGeometry.RowSize(sheet, range.Top) : (DrawingGeometry.ColumnSize(sheet, range.Left) - 5) / 7;
        var dialog = new FieldsDialog(row ? "行高" : "列宽").Text("size", row ? "行高（像素）" : "列宽（字符单位）", size.ToString(CultureInfo.CurrentCulture));
        dialog.Validate = () => dialog.Number("size", row ? 1 : 1, row ? 10000 : 1427);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetAxisMetadataCommand(sheet.Id, axis, row ? range.Top : range.Left, row ? range.Rows : range.Columns, dialog.Number("size", 1, row ? 10000 : 1427)), revision);
    }
}
