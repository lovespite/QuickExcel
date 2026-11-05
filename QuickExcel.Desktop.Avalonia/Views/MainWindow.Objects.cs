using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallObjects()
    {
        _grid.DrawingObjectChanged += change => _ = PerformAsync(async () =>
        {
            if (Document.Snapshot.DocumentId != change.DocumentId || Document.Snapshot.ActiveSheetId != change.SheetId) return;
            await ExecuteAsync(new TransformDrawingObjectCommand(change.SheetId, change.Target, change.Original, change.Placement), change.Revision);
        });
        _grid.DrawingObjectSelectionChanged += () =>
        {
            if (_chartFilterContextCurrent?.Invoke() == false) CloseChartFilter();
            UpdateFormula();
            if (_grid.SelectedDrawingObject is { } target) _status.Text = UiText.Format($"已选择{(target.Kind == DrawingObjectKind.Image ? UiText.Get("图片") : UiText.Get("图表"))} {target.Index + 1}：拖动移动，边框控制点缩放，顶部圆点旋转；Shift 保持比例或按 15° 旋转，Esc 取消。");
        };
        _grid.ChartFilterRequested += () => _ = PerformAsync(ChartFilterAsync);
        _grid.ViewportChanged += CloseChartFilter;
        RibbonGroup(UiText.Get("插入"), UiText.Get("对象"), RibbonButton(RibbonIcon.Images, UiText.Get("图片管理"), ImagesAsync), RibbonButton(RibbonIcon.Charts, UiText.Get("图表管理"), ChartsAsync), RibbonButton(RibbonIcon.ChartFilter, UiText.Get("图表筛选"), ChartFilterAsync), RibbonButton(RibbonIcon.Comment, UiText.Get("批注"), NoteAsync));
        RibbonGroup(UiText.Get("插入"), UiText.Get("单元格内容"), RibbonButton(RibbonIcon.PlainText, UiText.Get("富文本编辑"), RichTextAsync), RibbonButton(RibbonIcon.Comment, UiText.Get("超链接"), HyperlinkAsync), RibbonButton(RibbonIcon.Comment, UiText.Get("线程评论"), ThreadCommentsAsync));
        RibbonGroup(UiText.Get("页面布局"), UiText.Get("页面设置"), RibbonButton(RibbonIcon.PageSettings, UiText.Get("页面设置"), PageSettingsAsync), RibbonButton(RibbonIcon.SetPrintArea, UiText.Get("设为打印区域"), () => PrintAreaAsync(false)),
            RibbonButton(RibbonIcon.ClearPrintArea, UiText.Get("清除打印区域"), () => PrintAreaAsync(true)), RibbonButton(RibbonIcon.PageBreaks, UiText.Get("分页"), PageBreaksAsync));
    }
    private async Task NoteAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var position = _grid.Current; var old = sheet.Annotations.Notes.GetValueOrDefault(position);
        var dialog = new FieldsDialog(UiText.Get("传统批注")).Text("author", UiText.Get("作者"), old?.Author ?? "QuickExcel").Text("text", UiText.Get("批注内容"), old?.Text.Text ?? "", true)
            .Check("flatten", UiText.Get("替换含格式的批注时，明确转换为普通文本")).Check("remove", UiText.Get("删除此批注"));
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetNoteCommand(sheet.Id, position, dialog.Boolean("remove") ? null :
            new(dialog.String("author"), new([new TextRun(dialog.String("text"))])), dialog.Boolean("flatten")), snapshot.Revision);
    }
    private async Task ImagesAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var current = _grid.Current;
        var choose = new FieldsDialog(UiText.Get("图片管理")).Choice("image", UiText.Get("图片"), new[] { UiText.Get("插入新图片") }.Concat(sheet.Images.Select((_, i) => UiText.Format($"图片 {i + 1}"))).ToArray()).Check("remove", UiText.Get("删除所选图片"));
        if (_grid.SelectedDrawingObject is { Kind: DrawingObjectKind.Image } selected) ((ComboBox)choose.Input("image")).SelectedIndex = selected.Index + 1;
        if (!await choose.ShowAsync(this)) return; var index = choose.Index("image") - 1;
        if (choose.Boolean("remove")) { if (index < 0) throw new ArgumentException(UiText.Get("请先选择已有图片。")); await ExecuteAsync(new SetImageCommand(sheet.Id, index, null), snapshot.Revision); return; }
        SheetImage image;
        if (index < 0)
        {
            var file = await _files.OpenAsync(new(UiText.Get("插入图片"), [new(UiText.Get("图片"), ["*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.tif", "*.tiff"])])); if (file is null) return;
            var bytes = await Task.Run(() =>
            {
                using var input = new FileStream(file.Path, FileMode.Open, FileAccess.Read, FileShare.Read);
                if (input.Length > 16 * 1024 * 1024) throw new ArgumentException(UiText.Get("单张图片最多 16 MiB。"));
                var data = new byte[checked((int)input.Length)]; input.ReadExactly(data); return data;
            });
            image = new(bytes, Path.GetExtension(file.Path), current.Row, current.Column); index = sheet.Images.Length;
            using var renderer = new SkiaRenderer(); using var pixels = renderer.DecodeImage(image);
            image = image.WithAnchor(DrawingGeometry.FromBounds(sheet, new(DrawingGeometry.ColumnStart(sheet, current.Column), DrawingGeometry.RowStart(sheet, current.Row),
                Math.Min(pixels.Width, 1000), Math.Min(pixels.Width, 1000) * (double)pixels.Height / pixels.Width), DrawingAnchorKind.OneCell));
        }
        else image = sheet.Images[index];
        var bounds = image.Anchor.Bounds(sheet);
        var dialog = DrawingFields(UiText.Get("图片位置和尺寸"), bounds, image.Anchor.Kind, image.Anchor.Rotation);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetImageCommand(sheet.Id, index, image.WithAnchor(ReadAnchor(dialog, sheet))), snapshot.Revision);
    }
    private static FieldsDialog DrawingFields(string title, DrawingBounds bounds, DrawingAnchorKind kind, double rotation = 0)
    {
        var dialog = new FieldsDialog(title).Note(UiText.Get("位置和尺寸使用 96 DPI 像素。"))
            .Text("left", UiText.Get("距工作表左边"), bounds.X.ToString()).Text("top", UiText.Get("距工作表顶部"), bounds.Y.ToString())
            .Text("width", UiText.Get("宽度"), bounds.Width.ToString()).Text("height", UiText.Get("高度"), bounds.Height.ToString())
            .Text("rotation", UiText.Get("旋转角度（度）"), rotation.ToString())
            .Choice("behavior", UiText.Get("随单元格变化"), [UiText.Get("移动并缩放"), UiText.Get("移动但不缩放"), UiText.Get("固定位置和尺寸")], (int)kind);
        dialog.Validate = () => { ReadBounds(dialog); dialog.Number("rotation", -36000, 36000); }; return dialog;
    }
    private static DrawingBounds ReadBounds(FieldsDialog dialog) => new(dialog.Number("left", 0, 2_000_000_000), dialog.Number("top", 0, 2_000_000_000), dialog.Number("width", 1, 100000), dialog.Number("height", 1, 100000));
    private static DrawingAnchor ReadAnchor(FieldsDialog dialog, WorksheetSnapshot sheet) => DrawingGeometry.FromBounds(sheet, ReadBounds(dialog), (DrawingAnchorKind)dialog.Index("behavior"), DrawingGeometry.NormalizeRotation(Math.Round(dialog.Number("rotation", -36000, 36000) * 60000) / 60000));
    private async Task ChartsAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var selection = Rectangle(SelectedRange(allowObject: true));
        var dialog = new FieldsDialog(UiText.Get("图表管理")).Choice("chart", UiText.Get("图表"), new[] { UiText.Get("插入新图表") }.Concat(sheet.Data.Charts.Select((c, i) => $"{i + 1} · {c.Title}")).ToArray())
            .Choice("kind", UiText.Get("类型"), [UiText.Get("簇状柱形图"), UiText.Get("折线图"), UiText.Get("饼图"), UiText.Get("散点图"), UiText.Get("平滑曲线图"), UiText.Get("瀑布图"), UiText.Get("漏斗图")]).Text("title", UiText.Get("标题")).Text("range", UiText.Get("数据范围（首行标题，首列分类；编辑时留空保持）"), selection.ToString())
            .Note(UiText.Get("散点图首列为数值 X，其余列为 Y。饼图、瀑布图和漏斗图使用两列。"))
            .Text("totals", UiText.Get("瀑布图总计点（从 1 开始，用逗号分隔）"))
            .Choice("legend", UiText.Get("图例"), [UiText.Get("右侧"), UiText.Get("底部"), UiText.Get("左侧"), UiText.Get("顶部"), UiText.Get("隐藏")]).Check("position", UiText.Get("调整位置和尺寸")).Check("remove", UiText.Get("删除所选图表"));
        ((ComboBox)dialog.Input("chart")).SelectionChanged += (_, _) =>
        {
            if (dialog.Index("chart") == 0) return;
            var chart = sheet.Data.Charts[dialog.Index("chart") - 1]; ((TextBox)dialog.Input("title")).Text = chart.Title; ((TextBox)dialog.Input("range")).Text = "";
            ((TextBox)dialog.Input("totals")).Text = string.Join(",", chart.Series[0].TotalPoints.Select(i => i + 1));
            ((ComboBox)dialog.Input("kind")).SelectedIndex = (int)chart.Kind; ((ComboBox)dialog.Input("legend")).SelectedIndex = chart.Legend switch { "r" => 0, "b" => 1, "l" => 2, "t" => 3, _ => 4 };
        };
        if (_grid.SelectedDrawingObject is { Kind: DrawingObjectKind.Chart } selected) ((ComboBox)dialog.Input("chart")).SelectedIndex = selected.Index + 1;
        if (!await dialog.ShowAsync(this)) return;
        var index = dialog.Index("chart") == 0 ? sheet.Data.Charts.Length : dialog.Index("chart") - 1;
        if (dialog.Boolean("remove")) { await ExecuteAsync(new SetChartCommand(sheet.Id, index, null), snapshot.Revision); return; }
        var old = index < sheet.Data.Charts.Length ? sheet.Data.Charts[index] : null;
        var anchor = old?.Anchor ?? DrawingGeometry.FromBounds(sheet, new(DrawingGeometry.ColumnStart(sheet, selection.Left), DrawingGeometry.RowStart(sheet, Math.Min(WorkbookLimits.MaxRows - 1, selection.Bottom + 2)), 560, 320), DrawingAnchorKind.OneCell);
        var chart = string.IsNullOrWhiteSpace(dialog.String("range")) && old is not null ? SetChartCommand.ChangeKind(old, (WorksheetChartKind)dialog.Index("kind")) with { Title = dialog.String("title") } :
            SetChartCommand.FromRange(sheet, CellRectangle.Parse(dialog.String("range")), (WorksheetChartKind)dialog.Index("kind"), anchor, dialog.String("title"));
        if (old is not null && !string.IsNullOrWhiteSpace(dialog.String("range")) && old.Kind == chart.Kind)
            chart = old with { Title = chart.Title, Series = [.. chart.Series.Select((s, i) => i >= old.Series.Length ? s : old.Series[i] with {
                Name = s.Name, NameReference = s.NameReference, Categories = s.Categories, Values = s.Values,
                HiddenPoints = [], IsFiltered = false,
                TotalPoints = [.. old.Series[i].TotalPoints.Where(p => p < WorksheetChart.Count(WorksheetChart.ParseReference(s.Values)))],
                PointColors = old.Series[i].PointColors.IsEmpty ? [] : [.. Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Select(p => p < old.Series[i].PointColors.Length ? old.Series[i].PointColors[p] : old.Kind == WorksheetChartKind.Waterfall ? "" : old.Series[i].Color)] })] };
        if (chart.Kind == WorksheetChartKind.Waterfall)
            chart = chart with { Series = [chart.Series[0] with { TotalPoints = [.. dialog.String("totals").Split([',', '，'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(s => int.Parse(s) - 1)] }] };
        chart = chart with { Legend = dialog.Index("legend") switch { 0 => "r", 1 => "b", 2 => "l", 3 => "t", _ => null } };
        if (dialog.Boolean("position"))
        { var position = DrawingFields(UiText.Get("图表位置和尺寸"), chart.Anchor.Bounds(sheet), chart.Anchor.Kind, chart.Anchor.Rotation); if (!await position.ShowAsync(this)) return; chart = chart with { Anchor = ReadAnchor(position, sheet) }; }
        await ExecuteAsync(new SetChartCommand(sheet.Id, index, chart), snapshot.Revision);
    }
}
