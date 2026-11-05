using QuickExcel.Core;
using QuickExcel.Domain;
using SkiaSharp;
using UglyToad.PdfPig;

namespace QuickExcel.Rendering.Skia;

/// <summary>One immutable print task, shared by preview/PDF/platform printing. No calculation or worker dependencies.</summary>
public sealed class SkiaReport : IDisposable
{
    public PrintJob Job { get; }
    public SkiaRenderer Renderer { get; }
    private readonly object _gate = new();
    private readonly Dictionary<SheetImage, SKBitmap> _images = [];
    private long _imageBytes;
    private readonly CellRectangle[] _merges;
    public SkiaReport(WorkbookSnapshot book, WorksheetSnapshot sheet, string file,
        Func<WorkbookSnapshot, WorksheetSnapshot, int, int, StyleDefinition>? style = null,
        IEnumerable<string>? fontFiles = null, bool systemFonts = true, DateTime? timestamp = null,
        RenderBudget? budget = null, CancellationToken cancellationToken = default,
        Func<WorkbookSnapshot, WorksheetSnapshot, int, int, ConditionalCellVisual?>? visual = null)
    {
        Renderer = new(fontFiles, systemFonts, budget);
        try
        {
            var warnings = new List<string>(); var time = timestamp ?? DateTime.Now;
            foreach (var diagnostic in book.FileState?.CompatibilityDiagnostics ?? [])
                if ((diagnostic.SheetName is null || diagnostic.SheetName == sheet.Name) &&
                    (diagnostic.Feature.StartsWith("chart/", StringComparison.Ordinal) || diagnostic.Feature.StartsWith("drawing/", StringComparison.Ordinal)))
                    warnings.Add("图表/绘图仅输出已支持的部分：" + diagnostic.Message);
            var visible = PrintingLayout.VisibleSheet(book, sheet).WithPrinting(PrintJob.PreparePrinting(sheet.Printing, file, time, warnings));
            var images = new List<SheetImage>(); var charts = new List<WorksheetChart>();
            for (var i = 0; i < sheet.Images.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try { using var image = Renderer.DecodeImage(sheet.Images[i]); images.Add(sheet.Images[i]); }
                catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentException)
                { warnings.Add($"图片 {i + 1} 无法解码或绘制，本次跳过：{ex.Message}"); }
            }
            for (var i = 0; i < sheet.Data.Charts.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested(); var chart = sheet.Data.Charts[i];
                try
                {
                    var b = chart.Anchor.Bounds(visible);
                    using var recorder = new SKPictureRecorder();
                    var canvas = recorder.BeginRecording(SKRect.Create((float)b.X, (float)b.Y, (float)b.Width, (float)b.Height));
                    Renderer.DrawChart(canvas, new(b.X, b.Y, b.Width, b.Height), book, chart);
                    using var picture = recorder.EndRecording(); charts.Add(chart);
                }
                catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException or ArgumentException or FormatException)
                { warnings.Add($"图表 {i + 1}（{chart.Title}）无法绘制，本次跳过：{ex.Message}"); }
            }
            visible = visible.WithImages(images).WithData(visible.Data with { Charts = [.. charts] });
            Job = new(book, visible, file, time, warnings, style, cancellationToken, visual);
            if (Job.Layout.Pages.Any(p => (long)p.Rows.Length * p.Columns.Length > Renderer.Budget.MaxPageCells))
                throw new NotSupportedException($"单页超过 {Renderer.Budget.MaxPageCells:N0} 格渲染预算，请调整打印缩放或区域。");
            _merges = visible.MergedCells.Select(CellRectangle.Parse).ToArray();
        }
        catch { Renderer.Dispose(); throw; }
    }

    public void DrawPage(SKCanvas canvas, int pageNumber, CancellationToken cancellationToken = default)
    { lock (_gate) DrawPageCore(canvas, pageNumber, cancellationToken); }
    private void DrawPageCore(SKCanvas canvas, int pageNumber, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var sheet = Job.Sheet; var layout = Job.Layout; var page = layout.Pages[pageNumber]; var margins = layout.Margins; var options = sheet.Printing.Options;
        var monochrome = sheet.Printing.Setup?.BlackAndWhite == true;
        canvas.Clear(SKColors.White);
        var rows = page.Rows.Select(r => DrawingGeometry.RowSize(sheet, r)).ToArray(); var columns = page.Columns.Select(c => DrawingGeometry.ColumnSize(sheet, c)).ToArray();
        var headingWidth = options?.Headings == true ? 40 : 0; var headingHeight = options?.Headings == true ? 24 : 0;
        var width = columns.Sum() + headingWidth; var height = rows.Sum() + headingHeight;
        var x = margins.Left * 96; var y = margins.Top * 96;
        var usableWidth = layout.PaperWidth - (margins.Left + margins.Right) * 96; var usableHeight = layout.PaperHeight - (margins.Top + margins.Bottom) * 96;
        if (options?.HorizontalCentered == true) x += Math.Max(0, (usableWidth - width * layout.Scale) / 2);
        if (options?.VerticalCentered == true) y += Math.Max(0, (usableHeight - height * layout.Scale) / 2);
        canvas.Save();
        try
        {
            canvas.ClipRect(SKRect.Create((float)(margins.Left * 96), (float)(margins.Top * 96), (float)usableWidth, (float)usableHeight));
            canvas.Translate((float)x, (float)y); canvas.Scale((float)layout.Scale);
            if (options?.Headings == true)
            {
                var offset = 24d;
                for (var i = 0; i < rows.Length; i++) { Renderer.DrawText(canvas, new((page.Rows[i] + 1).ToString(), new()), SKRect.Create(0, (float)offset, 38, (float)rows[i])); offset += rows[i]; }
                offset = 40;
                for (var i = 0; i < columns.Length; i++)
                {
                    var label = QuickExcel.Formulas.FormulaSyntax.Print(new QuickExcel.Formulas.ReferenceNode(null, 0, page.Columns[i], 0, page.Columns[i]))[..^1];
                    Renderer.DrawText(canvas, new(label, new()), SKRect.Create((float)offset, 0, (float)columns[i], 22)); offset += columns[i];
                }
            }
            canvas.Translate(headingWidth, headingHeight);
            long textUnits = 0;
            var borders = new List<(RenderRect Bounds, StyleDefinition Style)>();
            void Cell(int row, int column, RenderRect rect)
            {
                token.ThrowIfCancellationRequested(); var input = Job.Cell(row, column, rect);
                textUnits += input.Text.Text.Length;
                if (textUnits > 2_000_000) throw new NotSupportedException("单页文字超过 2,000,000 字符渲染预算。");
                var merge = _merges.FirstOrDefault(m => m.Top == row && m.Left == column);
                borders.Add((rect, merge != default ? CellBorderProjection.ForDiagonals(input.Text.Style) : input.Text.Style));
                Renderer.DrawCell(canvas, input, borders: false, monochrome: monochrome, location: $"{sheet.Name}!R{row + 1}C{column + 1}");
            }
            foreach (var rb in Bands(page.Rows, rows))
            foreach (var cb in Bands(page.Columns, columns))
            {
                canvas.Save();
                try
                {
                    canvas.ClipRect(SKRect.Create((float)cb.Offset, (float)rb.Offset, (float)cb.Size, (float)rb.Size));
                    canvas.Translate((float)(cb.Offset - DrawingGeometry.ColumnStart(sheet, cb.First)), (float)(rb.Offset - DrawingGeometry.RowStart(sheet, rb.First)));
                    borders.Clear();
                    var merges = _merges.Where(m => m.Bottom >= rb.First && m.Top <= rb.Last && m.Right >= cb.First && m.Left <= cb.Last).ToArray();
                    var top = DrawingGeometry.RowStart(sheet, rb.First);
                    for (var r = rb.First; r <= rb.Last; r++)
                    {
                        var h = DrawingGeometry.RowSize(sheet, r); var left = DrawingGeometry.ColumnStart(sheet, cb.First);
                        for (var c = cb.First; c <= cb.Last; c++)
                        {
                            var w = DrawingGeometry.ColumnSize(sheet, c);
                            if (w > 0 && h > 0 && !merges.Any(m => m.Contains(r, c))) Cell(r, c, new(left, top, w, h)); left += w;
                        }
                        top += h;
                    }
                    foreach (var merge in merges)
                    {
                        Cell(merge.Top, merge.Left, new(DrawingGeometry.ColumnStart(sheet, merge.Left), DrawingGeometry.RowStart(sheet, merge.Top),
                            DrawingGeometry.ColumnStart(sheet, merge.Right + 1) - DrawingGeometry.ColumnStart(sheet, merge.Left), DrawingGeometry.RowStart(sheet, merge.Bottom + 1) - DrawingGeometry.RowStart(sheet, merge.Top)));
                        var left = DrawingGeometry.ColumnStart(sheet, merge.Left); var right = DrawingGeometry.ColumnStart(sheet, merge.Right + 1);
                        var mergeTop = DrawingGeometry.RowStart(sheet, merge.Top); var bottom = DrawingGeometry.RowStart(sheet, merge.Bottom + 1);
                        StyleDefinition StyleAt(int r, int c) => Job.Cell(r, c, default).Text.Style;
                        for (var r = Math.Max(merge.Top, rb.First); r <= Math.Min(merge.Bottom, rb.Last); r++)
                            if (DrawingGeometry.RowSize(sheet, r) is var h && h > 0)
                                borders.Add((new(left, DrawingGeometry.RowStart(sheet, r), right - left, h), CellBorderProjection.ForRow(merge, r, StyleAt)));
                        for (var c = Math.Max(merge.Left, cb.First); c <= Math.Min(merge.Right, cb.Last); c++)
                            if (DrawingGeometry.ColumnSize(sheet, c) is var w && w > 0)
                                borders.Add((new(DrawingGeometry.ColumnStart(sheet, c), mergeTop, w, bottom - mergeTop), CellBorderProjection.ForColumn(merge, c, StyleAt)));
                    }
                    foreach (var input in borders) SkiaRenderer.DrawBorders(canvas, input.Bounds, input.Style, monochrome);
                    var band = SKRect.Create((float)DrawingGeometry.ColumnStart(sheet, cb.First), (float)DrawingGeometry.RowStart(sheet, rb.First), (float)cb.Size, (float)rb.Size);
                    foreach (var image in sheet.Images)
                    {
                        token.ThrowIfCancellationRequested(); var b = image.Anchor.Bounds(sheet); var rect = SKRect.Create((float)b.X, (float)b.Y, (float)b.Width, (float)b.Height);
                        var rotated = DrawingGeometry.RotatedBounds(b, image.Anchor.Rotation);
                        if (!SKRect.Create((float)rotated.X, (float)rotated.Y, (float)rotated.Width, (float)rotated.Height).IntersectsWith(band)) continue;
                        if (!_images.TryGetValue(image, out var bitmap))
                        {
                            bitmap = Renderer.DecodeImage(image); var bytes = (long)bitmap.RowBytes * bitmap.Height;
                            if (_imageBytes + bytes > Renderer.Budget.MaxImageCacheBytes) ClearImages();
                            if (bytes > Renderer.Budget.MaxImageCacheBytes) { using (bitmap) SkiaRenderer.DrawImage(canvas, bitmap, rect, monochrome, image.Anchor.Rotation); continue; }
                            _images.Add(image, bitmap); _imageBytes += bytes;
                        }
                        SkiaRenderer.DrawImage(canvas, bitmap, rect, monochrome, image.Anchor.Rotation);
                    }
                    foreach (var chart in sheet.Data.Charts)
                    {
                        token.ThrowIfCancellationRequested(); var b = chart.Anchor.Bounds(sheet);
                        var rotated = DrawingGeometry.RotatedBounds(b, chart.Anchor.Rotation);
                        if (SKRect.Create((float)rotated.X, (float)rotated.Y, (float)rotated.Width, (float)rotated.Height).IntersectsWith(band)) Renderer.DrawChart(canvas, new(b.X, b.Y, b.Width, b.Height), Job.Book, chart, monochrome);
                    }
                }
                finally { canvas.Restore(); }
            }
        }
        finally { canvas.Restore(); }
        var hf = sheet.Printing.HeaderFooter; var number = pageNumber + (int)(sheet.Printing.Setup?.UseFirstPageNumber == true ? sheet.Printing.Setup.FirstPageNumber ?? 1 : 1);
        var first = pageNumber == 0 && hf?.DifferentFirst == true; var even = number % 2 == 0 && hf?.DifferentOddEven == true;
        Header(canvas, first ? hf?.FirstHeader : even ? hf?.EvenHeader : hf?.OddHeader, margins.Header * 96, number);
        Header(canvas, first ? hf?.FirstFooter : even ? hf?.EvenFooter : hf?.OddFooter, layout.PaperHeight - margins.Footer * 96 - 16, number);
    }
    private static IEnumerable<(int First, int Last, double Offset, double Size)> Bands(IReadOnlyList<int> indices, double[] sizes)
    {
        var offset = 0d;
        for (var i = 0; i < indices.Count;)
        {
            var first = i; var size = sizes[i++]; while (i < indices.Count && indices[i] == indices[i - 1] + 1) size += sizes[i++];
            yield return (indices[first], indices[i - 1], offset, size); offset += size;
        }
    }
    private void Header(SKCanvas canvas, string? source, double y, int number)
    {
        var text = PrintingLayout.Header(source, Job.Sheet.Name, Job.FileName, number, Job.Layout.Pages.Length, Job.Timestamp);
        var align = Job.Sheet.Printing.HeaderFooter?.AlignWithMargins != false;
        var left = align ? Job.Layout.Margins.Left * 96 : 24; var right = align ? Job.Layout.Margins.Right * 96 : 24;
        var scale = Job.Sheet.Printing.HeaderFooter?.ScaleWithDoc != false ? Job.Layout.Scale : 1;
        var rect = SKRect.Create(0, 0, (float)((Job.Layout.PaperWidth - left - right) / scale), 20);
        canvas.Save(); canvas.Translate((float)left, (float)y); canvas.Scale((float)scale);
        try
        {
            Renderer.DrawText(canvas, new(text.Left, new() { Horizontal = 1, FontSize = 10 }), rect);
            Renderer.DrawText(canvas, new(text.Center, new() { Horizontal = 2, FontSize = 10 }), rect);
            Renderer.DrawText(canvas, new(text.Right, new() { Horizontal = 3, FontSize = 10 }), rect);
        }
        finally { canvas.Restore(); }
    }
    public SKBitmap RenderPage(int pageNumber, double dpi = 96, CancellationToken cancellationToken = default)
    {
        if (!double.IsFinite(dpi) || dpi <= 0) throw new ArgumentOutOfRangeException(nameof(dpi));
        var width = checked((int)Math.Ceiling(Job.Layout.PaperWidth * dpi / 96)); var height = checked((int)Math.Ceiling(Job.Layout.PaperHeight * dpi / 96));
        if ((long)width * height * 4 > Renderer.Budget.MaxPreviewBytes) throw new NotSupportedException("预览像素超过缓存预算，请降低预览缩放。");
        var bitmap = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
        try { using var canvas = new SKCanvas(bitmap); canvas.Scale((float)(dpi / 96)); DrawPage(canvas, pageNumber, cancellationToken); return bitmap; }
        catch { bitmap.Dispose(); throw; }
    }
    public void WritePdf(string destination, CancellationToken cancellationToken = default)
    {
        lock (_gate)
        {
            var path = Path.GetFullPath(destination); var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                {
                    using var output = new PdfOutputStream(stream, Renderer.Budget.MaxPdfBytes);
                    // Skia rounds the media box to this grid. 600 DPI keeps the physical error below .12 pt.
                    using var pdf = SKDocument.CreatePdf(output, new SKDocumentPdfMetadata { Title = Job.FileName, Creator = "QuickExcel", Creation = Job.Timestamp, Modified = Job.Timestamp, RasterDpi = 600 });
                    output.ThrowIfFailed();
                    if (pdf is null) throw new IOException("Skia PDF backend is unavailable.");
                    for (var i = 0; i < Job.Layout.Pages.Length; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var canvas = pdf.BeginPage((float)(Job.Layout.PaperWidth * .75), (float)(Job.Layout.PaperHeight * .75));
                        canvas.Scale(.75f); DrawPageCore(canvas, i, cancellationToken); pdf.EndPage();
                        output.ThrowIfFailed();
                    }
                    pdf.Close(); output.ThrowIfFailed(); stream.Flush(true);
                }
                using (var parsed = PdfDocument.Open(temporary))
                {
                    if (parsed.NumberOfPages != Job.Layout.Pages.Length) throw new InvalidDataException("PDF 页数不匹配。");
                    foreach (var page in parsed.GetPages())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (Math.Abs(page.Width - Job.Layout.PaperWidth * .75) > .2 || Math.Abs(page.Height - Job.Layout.PaperHeight * .75) > .2)
                            throw new InvalidDataException($"PDF 纸张尺寸不匹配：{page.Width} × {page.Height}，预期 {Job.Layout.PaperWidth * .75} × {Job.Layout.PaperHeight * .75} pt。");
                        _ = page.Letters.Count; _ = page.NumberOfImages;
                    }
                }
                cancellationToken.ThrowIfCancellationRequested();
                File.Move(temporary, path, true);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
    private void ClearImages() { foreach (var image in _images.Values) image.Dispose(); _images.Clear(); _imageBytes = 0; }
    public void ReleaseDrawingCaches() { lock (_gate) { ClearImages(); Renderer.Fonts.ReleaseShapingCaches(); } }
    public void Dispose() { lock (_gate) { ClearImages(); Renderer.Dispose(); } }
}
