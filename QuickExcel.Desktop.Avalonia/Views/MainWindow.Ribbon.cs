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
    private readonly Dictionary<string, RibbonPage> _ribbonPages = [];
    private bool _updatingRibbon;
    private WorkbookClipboard? _formatBrush;
    private Guid _formatBrushSession;
    private StyleDefinition CurrentStyle => Document.Snapshot.Styles[Document.Snapshot.Sheets[SheetIndex].GetStyleId(_grid.Current.Row, _grid.Current.Column)];
    private Button RibbonButton(RibbonIcon icon, string title, Func<Task> action, bool mutation = true)
    {
        var button = new Button { Content = new RibbonCommandContent(icon, title), Padding = new(6, 3), Margin = new(2, 1), Height = 28, MinHeight = 28, MinWidth = 32,
            Background = Brushes.Transparent, BorderThickness = new(0),
            HorizontalAlignment = HorizontalAlignment.Left, HorizontalContentAlignment = HorizontalAlignment.Left };
        button.Classes.Add("ribbon-command");
        AutomationProperties.SetName(button, title); ToolTip.SetTip(button, title);
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
        if (name == UiText.Get("字体"))
        {
            body.Children.Clear(); body.MinWidth = 236; body.ColumnDefinitions = new("*");
            var selectors = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            selectors.Children.Add(_font); selectors.Children.Add(_fontSize);
            body.Children.Add(selectors);
            var formatting = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
            for (var i = 2; i < controls.Length; i++)
            {
                formatting.Children.Add(controls[i]);
            }
            global::Avalonia.Controls.Grid.SetRow(formatting, 1);
            global::Avalonia.Controls.Grid.SetRowSpan(formatting, 2); body.Children.Add(formatting);
        }
        RibbonContentGroup(page, name, body);
    }
    private void RibbonContentGroup(string page, string name, Control body)
    {
        if (!_ribbonPages.TryGetValue(page, out var scroll))
        {
            scroll = new RibbonPage(allowLabels: page != UiText.Get("开始"));
            _ribbonPages.Add(page, scroll);
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
        scroll.AddGroup(new Border { BorderBrush = Brushes.LightGray, BorderThickness = new(0, 0, 1, 0), Child = panel });
    }
    private void ArrangeRibbonTabs()
    {
        string[] order = [UiText.Get("开始"), UiText.Get("插入"), UiText.Get("页面布局"), UiText.Get("公式"), UiText.Get("自动执行"), UiText.Get("数据"), UiText.Get("审阅"), UiText.Get("视图")];
        var tabs = _ribbon.Items.Cast<TabItem>().OrderBy(t => Array.IndexOf(order, t.Header as string)).ToArray();
        _ribbon.Items.Clear(); foreach (var tab in tabs) _ribbon.Items.Add(tab); _ribbon.SelectedIndex = 0;
    }
    private void InstallRibbon()
    {
        AutomationProperties.SetName(_ribbon, UiText.Get("功能区")); DockPanel.SetDock(_ribbon, Dock.Top); _book.Children.Insert(0, _ribbon);
        RibbonGroup(UiText.Get("开始"), UiText.Get("剪贴板"), RibbonButton(RibbonIcon.Paste, UiText.Get("粘贴"), () => PasteAsync("all")), RibbonButton(RibbonIcon.Copy, UiText.Get("复制"), () => CopyAsync(false), false),
            RibbonButton(RibbonIcon.Cut, UiText.Get("剪切"), () => CopyAsync(true)), RibbonButton(RibbonIcon.PasteValues, UiText.Get("只贴值"), () => PasteAsync("values")), RibbonButton(RibbonIcon.FormatBrush, UiText.Get("格式刷"), FormatBrushAsync));
        _font.ItemsSource = FontManager.Current.SystemFonts.Select(f => f.Name).Order(StringComparer.CurrentCultureIgnoreCase).ToArray();
        _fontSize.ItemsSource = new double[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72 };
        AutomationProperties.SetName(_font, UiText.Get("字体")); AutomationProperties.SetName(_fontSize, UiText.Get("字号"));
        _font.SelectionChanged += async (_, _) => { if (!_updatingRibbon && _font.SelectedItem is string name) await PerformAsync(() => ApplyStyleAsync(new(FontName: name))); };
        _fontSize.SelectionChanged += async (_, _) => { if (!_updatingRibbon && _fontSize.SelectedItem is double size) await PerformAsync(() => ApplyStyleAsync(new(FontSize: size))); };
        _mutationControls.AddRange([_font, _fontSize]);
        RibbonGroup(UiText.Get("开始"), UiText.Get("字体"), _font, _fontSize, RibbonButton(RibbonIcon.Bold, UiText.Get("加粗"), () => ApplyStyleAsync(new(Bold: !CurrentStyle.Bold))),
            RibbonButton(RibbonIcon.Italic, UiText.Get("倾斜"), () => ApplyStyleAsync(new(Italic: !CurrentStyle.Italic))), RibbonColorDropdown(false), RibbonColorDropdown(true), BorderDropdown());
        RibbonGroup(UiText.Get("开始"), UiText.Get("对齐"), RibbonButton(RibbonIcon.AlignLeft, UiText.Get("左对齐"), () => ApplyStyleAsync(new(Horizontal: 1))), RibbonButton(RibbonIcon.AlignCenter, UiText.Get("居中"), () => ApplyStyleAsync(new(Horizontal: 2))),
            RibbonButton(RibbonIcon.AlignRight, UiText.Get("右对齐"), () => ApplyStyleAsync(new(Horizontal: 3))),
            RibbonButton(RibbonIcon.AlignTop, UiText.Get("顶端对齐"), () => ApplyStyleAsync(new(Vertical: 0))),
            RibbonButton(RibbonIcon.AlignMiddle, UiText.Get("垂直居中"), () => ApplyStyleAsync(new(Vertical: 1))),
            RibbonButton(RibbonIcon.AlignBottom, UiText.Get("底端对齐"), () => ApplyStyleAsync(new(Vertical: 2))),
            TextRotationDropdown(),
            RibbonButton(RibbonIcon.WrapText, UiText.Get("自动换行"), () => ApplyStyleAsync(new(WrapText: !CurrentStyle.WrapText))),
            RibbonButton(RibbonIcon.MergeCells, UiText.Get("合并"), () => MergeAsync(false)), RibbonButton(RibbonIcon.SplitCells, UiText.Get("取消合并"), () => MergeAsync(true)));
        InstallNumberFormatRibbon();
        InstallCellStyles();
        RibbonGroup(UiText.Get("开始"), UiText.Get("单元格"), RibbonButton(RibbonIcon.FormatCells, UiText.Get("设置单元格格式"), CellFormatAsync), RibbonButton(RibbonIcon.ClearFormatting, UiText.Get("清除格式"), ClearFormatsAsync));
        RibbonGroup(UiText.Get("公式"), UiText.Get("公式"), RibbonButton(RibbonIcon.AutoSum, UiText.Get("自动求和  Alt+="), () => { BeginAutoSum(); return Task.CompletedTask; }), RibbonButton(RibbonIcon.Calculate, UiText.Get("重算  F9"), CalculateAsync, false));
        RibbonGroup(UiText.Get("开始"), UiText.Get("编辑"), RibbonButton(RibbonIcon.FillDown, UiText.Get("向下填充"), () => FillDirectionAsync(false)), RibbonButton(RibbonIcon.FillRight, UiText.Get("向右填充"), () => FillDirectionAsync(true)), RibbonButton(RibbonIcon.FillSeries, UiText.Get("序列"), SeriesAsync),
            RibbonButton(RibbonIcon.PlainText, UiText.Get("转为普通文本"), FlattenRichTextAsync));
        _grid.AutoFitRequested += (axis, index) => _ = PerformAsync(() => AutoFitAsync(axis, index));
        RibbonGroup(UiText.Get("视图"), UiText.Get("行列尺寸"), RibbonButton(RibbonIcon.AutoFitHeight, UiText.Get("自动行高"), () => AutoFitAsync(SheetAxis.Row, _grid.Current.Row)),
            RibbonButton(RibbonIcon.AutoFitWidth, UiText.Get("自动列宽"), () => AutoFitAsync(SheetAxis.Column, _grid.Current.Column)), RibbonButton(RibbonIcon.RowHeight, UiText.Get("行高"), () => AxisSizeAsync(SheetAxis.Row)), RibbonButton(RibbonIcon.ColumnWidth, UiText.Get("列宽"), () => AxisSizeAsync(SheetAxis.Column)));
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
            _numberFormatLabel.Text = NumberFormatOptions.Label(style);
            UpdateColorSwatches();
            UpdateCellStyleSelection();
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
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange();
        if (await ExecuteAsync(new MergeCellsCommand(Document.Snapshot.ActiveSheetId, Rectangle(range), remove)) && !remove)
            _grid.Select(new(range.Top, range.Left));
    }
    private async Task FormatBrushAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var snapshot = Document.Snapshot;
        if (_formatBrush is null)
        {
            _formatBrush = WorkbookClipboard.Capture(snapshot, snapshot.ActiveSheetId, new(_grid.Current.Row, _grid.Current.Column, _grid.Current.Row, _grid.Current.Column));
            _formatBrushSession = Document.Session.SessionId; _status.Text = UiText.Get("已采样格式，选择目标后再次点击格式刷；Esc 取消。"); return;
        }
        var brush = _formatBrush; _formatBrush = null;
        if (_formatBrushSession != Document.Session.SessionId) throw new InvalidOperationException(UiText.Get("原工作簿已关闭，请重新采样格式。"));
        var range = SelectedRange(); RequireVisiblePaste(range);
        await ExecuteAsync(new PasteCellsCommand(brush, snapshot.ActiveSheetId, Rectangle(range), brush.Revision, CellPasteMode.Formats), brush.Revision);
    }
    private async Task FlattenRichTextAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses();
        if (await Dialogs.Confirm(this, UiText.Get("转换为普通文本"), UiText.Get("将移除选中单元格的分段字体格式，保留文字内容。可撤销。")))
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
        var dialog = new FieldsDialog(row ? UiText.Get("行高") : UiText.Get("列宽")).Text("size", row ? UiText.Get("行高（像素）") : UiText.Get("列宽（字符单位）"), size.ToString(CultureInfo.CurrentCulture));
        dialog.Validate = () => dialog.Number("size", row ? 1 : 1, row ? 10000 : 1427);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetAxisMetadataCommand(sheet.Id, axis, row ? range.Top : range.Left, row ? range.Rows : range.Columns, dialog.Number("size", 1, row ? 10000 : 1427)), revision);
    }
}
