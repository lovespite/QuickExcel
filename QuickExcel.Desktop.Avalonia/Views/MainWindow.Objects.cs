using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallObjects()
    {
        RibbonGroup("插入", "对象", RibbonButton("图片管理…", ImagesAsync), RibbonButton("图表管理…", ChartsAsync), RibbonButton("批注…", NoteAsync));
        RibbonGroup("页面布局", "页面设置", RibbonButton("页面设置…", PageSettingsAsync), RibbonButton("设为打印区域", () => PrintAreaAsync(false)),
            RibbonButton("清除打印区域", () => PrintAreaAsync(true)), RibbonButton("分页…", PageBreaksAsync));
    }
    private async Task NoteAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var position = _grid.Current; var old = sheet.Annotations.Notes.GetValueOrDefault(position);
        var dialog = new FieldsDialog("传统批注").Text("author", "作者", old?.Author ?? "QuickExcel").Text("text", "批注内容", old?.Text.Text ?? "", true)
            .Check("flatten", "替换含格式的批注时，明确转换为普通文本").Check("remove", "删除此批注");
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetNoteCommand(sheet.Id, position, dialog.Boolean("remove") ? null :
            new(dialog.String("author"), new([new TextRun(dialog.String("text"))])), dialog.Boolean("flatten")), snapshot.Revision);
    }
    private async Task ImagesAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var current = _grid.Current;
        var choose = new FieldsDialog("图片管理").Choice("image", "图片", new[] { "插入新图片…" }.Concat(sheet.Images.Select((_, i) => $"图片 {i + 1}")).ToArray()).Check("remove", "删除所选图片");
        if (!await choose.ShowAsync(this)) return; var index = choose.Index("image") - 1;
        if (choose.Boolean("remove")) { if (index < 0) throw new ArgumentException("请先选择已有图片。"); await ExecuteAsync(new SetImageCommand(sheet.Id, index, null), snapshot.Revision); return; }
        SheetImage image;
        if (index < 0)
        {
            var file = await _files.OpenAsync(new("插入图片", [new("图片", ["*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.tif", "*.tiff"])])); if (file is null) return;
            var bytes = await Task.Run(() =>
            {
                using var input = new FileStream(file.Path, FileMode.Open, FileAccess.Read, FileShare.Read);
                if (input.Length > 16 * 1024 * 1024) throw new ArgumentException("单张图片最多 16 MiB。");
                var data = new byte[checked((int)input.Length)]; input.ReadExactly(data); return data;
            });
            image = new(bytes, Path.GetExtension(file.Path), current.Row, current.Column); index = sheet.Images.Length;
            using var renderer = new SkiaRenderer(); using var pixels = renderer.DecodeImage(image);
            image = image.WithAnchor(DrawingGeometry.FromBounds(sheet, new(DrawingGeometry.ColumnStart(sheet, current.Column), DrawingGeometry.RowStart(sheet, current.Row),
                Math.Min(pixels.Width, 1000), Math.Min(pixels.Width, 1000) * (double)pixels.Height / pixels.Width), DrawingAnchorKind.OneCell));
        }
        else image = sheet.Images[index];
        var bounds = image.Anchor.Bounds(sheet);
        var dialog = DrawingFields("图片位置和尺寸", bounds, image.Anchor.Kind);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetImageCommand(sheet.Id, index, image.WithAnchor(ReadAnchor(dialog, sheet))), snapshot.Revision);
    }
    private static FieldsDialog DrawingFields(string title, DrawingBounds bounds, DrawingAnchorKind kind)
    {
        var dialog = new FieldsDialog(title).Note("位置和尺寸使用 96 DPI 像素。")
            .Text("left", "距工作表左边", bounds.X.ToString()).Text("top", "距工作表顶部", bounds.Y.ToString())
            .Text("width", "宽度", bounds.Width.ToString()).Text("height", "高度", bounds.Height.ToString())
            .Choice("behavior", "随单元格变化", ["移动并缩放", "移动但不缩放", "固定位置和尺寸"], (int)kind);
        dialog.Validate = () => ReadBounds(dialog); return dialog;
    }
    private static DrawingBounds ReadBounds(FieldsDialog dialog) => new(dialog.Number("left", 0, 2_000_000_000), dialog.Number("top", 0, 2_000_000_000), dialog.Number("width", 1, 100000), dialog.Number("height", 1, 100000));
    private static DrawingAnchor ReadAnchor(FieldsDialog dialog, WorksheetSnapshot sheet) => DrawingGeometry.FromBounds(sheet, ReadBounds(dialog), (DrawingAnchorKind)dialog.Index("behavior"));
    private async Task ChartsAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var selection = Rectangle(SelectedRange());
        var dialog = new FieldsDialog("图表管理").Choice("chart", "图表", new[] { "插入新图表" }.Concat(sheet.Data.Charts.Select((c, i) => $"{i + 1} · {c.Title}")).ToArray())
            .Choice("kind", "类型", ["簇状柱形图", "折线图", "饼图"]).Text("title", "标题").Text("range", "数据范围（首行标题，首列分类；编辑时留空保持）", selection.ToString())
            .Choice("legend", "图例", ["右侧", "底部", "左侧", "顶部", "隐藏"]).Check("position", "调整位置和尺寸…").Check("remove", "删除所选图表");
        ((ComboBox)dialog.Input("chart")).SelectionChanged += (_, _) =>
        {
            if (dialog.Index("chart") == 0) return;
            var chart = sheet.Data.Charts[dialog.Index("chart") - 1]; ((TextBox)dialog.Input("title")).Text = chart.Title; ((TextBox)dialog.Input("range")).Text = "";
            ((ComboBox)dialog.Input("kind")).SelectedIndex = (int)chart.Kind; ((ComboBox)dialog.Input("legend")).SelectedIndex = chart.Legend switch { "r" => 0, "b" => 1, "l" => 2, "t" => 3, _ => 4 };
        };
        if (!await dialog.ShowAsync(this)) return;
        var index = dialog.Index("chart") == 0 ? sheet.Data.Charts.Length : dialog.Index("chart") - 1;
        if (dialog.Boolean("remove")) { await ExecuteAsync(new SetChartCommand(sheet.Id, index, null), snapshot.Revision); return; }
        var old = index < sheet.Data.Charts.Length ? sheet.Data.Charts[index] : null;
        var anchor = old?.Anchor ?? DrawingGeometry.FromBounds(sheet, new(DrawingGeometry.ColumnStart(sheet, selection.Left), DrawingGeometry.RowStart(sheet, Math.Min(WorkbookLimits.MaxRows - 1, selection.Bottom + 2)), 560, 320), DrawingAnchorKind.OneCell);
        var chart = string.IsNullOrWhiteSpace(dialog.String("range")) && old is not null ? old with { Kind = (WorksheetChartKind)dialog.Index("kind"), Title = dialog.String("title") } :
            SetChartCommand.FromRange(sheet, CellRectangle.Parse(dialog.String("range")), (WorksheetChartKind)dialog.Index("kind"), anchor, dialog.String("title"));
        chart = chart with { Legend = dialog.Index("legend") switch { 0 => "r", 1 => "b", 2 => "l", 3 => "t", _ => null } };
        if (dialog.Boolean("position"))
        { var position = DrawingFields("图表位置和尺寸", chart.Anchor.Bounds(sheet), chart.Anchor.Kind); if (!await position.ShowAsync(this)) return; chart = chart with { Anchor = ReadAnchor(position, sheet) }; }
        await ExecuteAsync(new SetChartCommand(sheet.Id, index, chart), snapshot.Revision);
    }
}
