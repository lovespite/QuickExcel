using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CopyDrawingPictureAsync()
    {
        if (!await CommitEditorsAsync()) return;
        if (_grid.SelectedDrawingObject is not { } selected) throw new InvalidOperationException("请先选择图片或图表对象。");
        await CopyDrawingAsync(selected, pictureOnly: true);
    }
    private async Task CopyDrawingAsync(DrawingObjectId target, bool pictureOnly = false)
    {
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        var png = RenderDrawingPng(snapshot, sheet, target);
        var capture = target.Kind == DrawingObjectKind.Chart && !pictureOnly ? ChartClipboard.Capture(snapshot, sheet.Data.Charts[target.Index]) : null;
        var drawing = capture is null ? null : OfficeDrawingClipboardCodec.WritePrivateChart(capture);
        var office = capture is null ? null : OfficeDrawingClipboardCodec.WriteOfficeChart(capture);
        await _clipboard.WriteAsync(new(null, Png: png, ChartPackage: drawing, OfficeDrawing: office));
        _clipboardCells = null; _clipboardToken = null;
        _status.Text = drawing is null ? "已复制对象图片。" : office is null ? "已复制图表和图片；此图表仅在 QuickExcel 内保留可编辑形式，对外粘贴图片。" : "已复制可编辑图表和图片；通过粘贴菜单选择形式。图表数据使用复制时的值快照，不连接源文件。";
    }
    private static byte[] RenderDrawingPng(WorkbookSnapshot book, WorksheetSnapshot sheet, DrawingObjectId target)
    {
        var anchor = target.Kind == DrawingObjectKind.Image ? sheet.Images[target.Index].Anchor : sheet.Data.Charts[target.Index].Anchor;
        var bounds = anchor.Bounds(sheet); var extent = DrawingGeometry.RotatedBounds(new(0, 0, bounds.Width, bounds.Height), anchor.Rotation);
        var width = checked((int)Math.Ceiling(extent.Width)); var height = checked((int)Math.Ceiling(extent.Height));
        if (width <= 0 || height <= 0 || (long)width * height > 16_777_216) throw new NotSupportedException("对象图片超过 1600 万像素，请先缩小对象。");
        using var renderer = new SkiaRenderer(); using var bitmap = new SKBitmap(width, height); using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.Transparent); canvas.Translate((float)-extent.X, (float)-extent.Y);
        if (target.Kind == DrawingObjectKind.Image)
        {
            using var source = renderer.DecodeImage(sheet.Images[target.Index]);
            SkiaRenderer.DrawImage(canvas, source, SKRect.Create(0, 0, (float)bounds.Width, (float)bounds.Height), false, anchor.Rotation);
        }
        else
        {
            canvas.RotateDegrees((float)anchor.Rotation, (float)bounds.Width / 2, (float)bounds.Height / 2);
            renderer.DrawChart(canvas, new(0, 0, bounds.Width, bounds.Height), book, sheet.Data.Charts[target.Index]);
        }
        using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100); var bytes = encoded.ToArray();
        if (bytes.Length > OfficeDrawingClipboardCodec.MaxBytes) throw new NotSupportedException("对象图片超过 16 MiB。"); return bytes;
    }
    private async Task PasteDrawingAsync(ClipboardExchange data, string mode, CellPosition target, long revision)
    {
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        RequireVisiblePaste(new(target.Row, target.Column, target.Row, target.Column));
        var chart = mode == "chart" || mode == "all" && (data.ChartPackage is not null || data.OfficeDrawing is { } office && OfficeDrawingClipboardCodec.ContainsChart(office));
        if (chart)
        {
            if (data.OfficeDrawing is null && data.ChartPackage is null) throw new NotSupportedException("剪贴板没有可编辑图表。若来源只提供图片，请选择粘贴为图片。");
            var source = data.ChartPackage is { } package ? OfficeDrawingClipboardCodec.ReadPrivateChart(package) : OfficeDrawingClipboardCodec.ReadChart(data.OfficeDrawing!);
            if (!await ExecuteAsync(new PasteChartCommand(source, sheet.Id, target), revision)) return;
            await CalculateAsync();
            _grid.SelectDrawingObject(new(DrawingObjectKind.Chart, sheet.Data.Charts.Length));
            _status.Text = "已粘贴可编辑图表，数据快照保存在新增的“图表数据”工作表；不连接源文件，可一次撤销。";
        }
        else
        {
            if (data.Png is null) throw new NotSupportedException("剪贴板没有位图；请在源应用中复制图片后重试。");
            if (data.Png.Length > OfficeDrawingClipboardCodec.MaxBytes) throw new InvalidDataException("图片超过 16 MiB。");
            if (data.Png.Length < 8 || !data.Png.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 })) throw new InvalidDataException("剪贴板 PNG 签名无效。");
            var image = new SheetImage(data.Png, "png", target.Row, target.Column);
            using var renderer = new SkiaRenderer(); using var bitmap = renderer.DecodeImage(image);
            image = image.WithAnchor(DrawingGeometry.FromBounds(sheet, new(DrawingGeometry.ColumnStart(sheet, target.Column), DrawingGeometry.RowStart(sheet, target.Row), bitmap.Width, bitmap.Height), DrawingAnchorKind.OneCell));
            if (!await ExecuteAsync(new SetImageCommand(sheet.Id, sheet.Images.Length, image), revision)) return;
            _grid.SelectDrawingObject(new(DrawingObjectKind.Image, sheet.Images.Length));
            _status.Text = "已粘贴图片，可移动、缩放、旋转，可一次撤销。";
        }
    }
}
