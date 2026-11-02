using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Core;
using QuickExcel.Calculation;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed class ReportDocument(WorkbookPaginator paginator) : IDocumentPaginatorSource
{
    public DocumentPaginator DocumentPaginator => paginator;
}

internal sealed class WorkbookPaginator : DocumentPaginator
{
    private readonly WorkbookSnapshot _book;
    private readonly WorksheetSnapshot _sheet;
    private readonly string _file;
    private readonly DateTime _timestamp = DateTime.Now;
    private readonly Dictionary<SheetImage, BitmapSource> _images = [];
    internal PrintLayout Layout { get; }
    internal WorkbookPaginator(WorkbookSnapshot book, WorksheetSnapshot sheet, string file)
    {
        if (book.FileState?.CompatibilityDiagnostics.Any(d => (d.SheetName is null || d.SheetName == sheet.Name) &&
            (d.Feature.StartsWith("chart/", StringComparison.Ordinal) || d.Feature.StartsWith("drawing/", StringComparison.Ordinal))) == true)
            throw new NotSupportedException("文件含未支持的图表或绘图设置，不能完整打印；请先查看兼容性报告并在 Excel 中处理。");
        _book = book; _sheet = PrintingLayout.VisibleSheet(book, sheet); _file = Path.GetFileName(file); Layout = PrintingLayout.Create(_sheet);
        foreach (var chart in sheet.Data.Charts) _ = ChartData.Read(book, chart);
        foreach (var image in sheet.Images)
        {
            using var stream = new MemoryStream(image.Data.ToArray(), false);
            var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.DelayCreation | BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.None);
            var bitmap = decoder.Frames[0];
            if ((long)bitmap.PixelWidth * bitmap.PixelHeight > 16_000_000) throw new NotSupportedException("图片解码超过 1600 万像素，不能完整打印。");
            BitmapSource loaded = BitmapFrame.Create(bitmap);
            if (sheet.Printing.Setup?.BlackAndWhite == true) loaded = new FormatConvertedBitmap(loaded, PixelFormats.Gray8, null, 0);
            var pixels = new byte[checked(loaded.PixelHeight * ((loaded.PixelWidth * loaded.Format.BitsPerPixel + 7) / 8))];
            var stride = (loaded.PixelWidth * loaded.Format.BitsPerPixel + 7) / 8; loaded.CopyPixels(pixels, stride, 0);
            loaded = BitmapSource.Create(loaded.PixelWidth, loaded.PixelHeight, loaded.DpiX, loaded.DpiY, loaded.Format, loaded.Palette, pixels, stride);
            loaded.Freeze(); _images.Add(image, loaded);
        }
        // Validate all header variants before opening the preview or sending any page to a printer.
        if (sheet.Printing.HeaderFooter is { } h)
            foreach (var text in new[] { h.OddHeader, h.EvenHeader, h.FirstHeader, h.OddFooter, h.EvenFooter, h.FirstFooter })
                _ = PrintingLayout.Header(text, sheet.Name, _file, 1, PageCount, _timestamp);
    }
    public override bool IsPageCountValid => true;
    public override int PageCount => Layout.Pages.Length;
    public override Size PageSize { get => new(Layout.PaperWidth, Layout.PaperHeight); set { if (value != PageSize) throw new NotSupportedException("请使用页面设置更改纸张，以保持预览和打印分页一致。"); } }
    public override IDocumentPaginatorSource Source => new ReportDocument(this);
    public override DocumentPage GetPage(int pageNumber)
    {
        if (pageNumber < 0 || pageNumber >= PageCount) return DocumentPage.Missing;
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen()) Paint(dc, pageNumber);
        return new(visual, PageSize, new Rect(PageSize), new Rect(PageSize));
    }
    internal void Paint(DrawingContext dc, int pageNumber)
    {
        var page = Layout.Pages[pageNumber]; var margins = Layout.Margins; var options = _sheet.Printing.Options;
        dc.DrawRectangle(Brushes.White, null, new Rect(PageSize));
        var rowSizes = page.Rows.Select(r => DrawingGeometry.RowSize(_sheet, r)).ToArray();
        var colSizes = page.Columns.Select(c => DrawingGeometry.ColumnSize(_sheet, c)).ToArray();
        var headings = options?.Headings == true; var headingWidth = headings ? 40 : 0; var headingHeight = headings ? 24 : 0;
        var contentWidth = colSizes.Sum() + headingWidth; var contentHeight = rowSizes.Sum() + headingHeight;
        var x = margins.Left * 96; var y = margins.Top * 96;
        if (options?.HorizontalCentered == true) x += Math.Max(0, (PageSize.Width - (margins.Left + margins.Right) * 96 - contentWidth * Layout.Scale) / 2);
        if (options?.VerticalCentered == true) y += Math.Max(0, (PageSize.Height - (margins.Top + margins.Bottom) * 96 - contentHeight * Layout.Scale) / 2);
        dc.PushClip(new RectangleGeometry(new Rect(margins.Left * 96, margins.Top * 96, PageSize.Width - (margins.Left + margins.Right) * 96, PageSize.Height - (margins.Top + margins.Bottom) * 96)));
        dc.PushTransform(new TranslateTransform(x, y)); dc.PushTransform(new ScaleTransform(Layout.Scale, Layout.Scale));
        if (headings)
        {
            var offset = 24d;
            for (var i = 0; i < page.Rows.Length; i++) { Text(dc, (page.Rows[i] + 1).ToString(), new Rect(0, offset, 38, rowSizes[i]), new()); offset += rowSizes[i]; }
            offset = 40;
            for (var i = 0; i < page.Columns.Length; i++) { Text(dc, QuickExcel.Formulas.FormulaSyntax.Print(new QuickExcel.Formulas.ReferenceNode(null, 0, page.Columns[i], 0, page.Columns[i]))[..^1], new Rect(offset, 0, colSizes[i], 22), new()); offset += colSizes[i]; }
        }
        dc.PushTransform(new TranslateTransform(headingWidth, headingHeight));
        foreach (var rb in Bands(page.Rows, rowSizes))
        foreach (var cb in Bands(page.Columns, colSizes))
        {
            dc.PushClip(new RectangleGeometry(new Rect(cb.Offset, rb.Offset, cb.Size, rb.Size)));
            dc.PushTransform(new TranslateTransform(cb.Offset - DrawingGeometry.ColumnStart(_sheet, cb.First), rb.Offset - DrawingGeometry.RowStart(_sheet, rb.First)));
            var merges = _sheet.MergedCells.Select(CellRectangle.Parse).Where(m => m.Bottom >= rb.First && m.Top <= rb.Last && m.Right >= cb.First && m.Left <= cb.Last).ToArray();
            var top = DrawingGeometry.RowStart(_sheet, rb.First);
            for (var r = rb.First; r <= rb.Last; r++)
            {
                var height = DrawingGeometry.RowSize(_sheet, r); var left = DrawingGeometry.ColumnStart(_sheet, cb.First);
                for (var c = cb.First; c <= cb.Last; c++)
                {
                    var width = DrawingGeometry.ColumnSize(_sheet, c);
                    if (width > 0 && height > 0 && !merges.Any(m => m.Contains(r, c))) Cell(dc, r, c, new(left, top, width, height));
                    left += width;
                }
                top += height;
            }
            foreach (var merge in merges)
                Cell(dc, merge.Top, merge.Left, new(DrawingGeometry.ColumnStart(_sheet, merge.Left), DrawingGeometry.RowStart(_sheet, merge.Top),
                    DrawingGeometry.ColumnStart(_sheet, merge.Right + 1) - DrawingGeometry.ColumnStart(_sheet, merge.Left), DrawingGeometry.RowStart(_sheet, merge.Bottom + 1) - DrawingGeometry.RowStart(_sheet, merge.Top)));
            foreach (var (picture, bitmap) in _images)
            {
                var b = picture.Anchor.Bounds(_sheet);
                dc.DrawImage(bitmap, new Rect(b.X, b.Y, b.Width, b.Height));
            }
            foreach (var chart in _sheet.Data.Charts)
            {
                var b = chart.Anchor.Bounds(_sheet); ChartDrawing.Paint(dc, new(b.X, b.Y, b.Width, b.Height), _book, chart, failIfStale: true, monochrome: _sheet.Printing.Setup?.BlackAndWhite == true);
            }
            dc.Pop(); dc.Pop();
        }
        dc.Pop(); dc.Pop(); dc.Pop(); dc.Pop();
        var hf = _sheet.Printing.HeaderFooter; var number = pageNumber + (int)(_sheet.Printing.Setup?.UseFirstPageNumber == true ? _sheet.Printing.Setup.FirstPageNumber ?? 1 : 1);
        var first = pageNumber == 0 && hf?.DifferentFirst == true; var even = number % 2 == 0 && hf?.DifferentOddEven == true;
        Header(dc, first ? hf?.FirstHeader : even ? hf?.EvenHeader : hf?.OddHeader, margins.Header * 96, number);
        Header(dc, first ? hf?.FirstFooter : even ? hf?.EvenFooter : hf?.OddFooter, PageSize.Height - margins.Footer * 96 - 16, number);
    }
    private static IEnumerable<(int First, int Last, double Offset, double Size)> Bands(IReadOnlyList<int> indices, double[] sizes)
    {
        var offset = 0d;
        for (var i = 0; i < indices.Count;)
        {
            var first = i; var size = sizes[i++];
            while (i < indices.Count && indices[i] == indices[i - 1] + 1) size += sizes[i++];
            yield return (indices[first], indices[i - 1], offset, size); offset += size;
        }
    }
    private void Cell(DrawingContext dc, int row, int column, Rect rect)
    {
        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0) return;
        var cell = _sheet.GetCell(row, column); var style = ConditionalPresentation.GetStyle(_book, _sheet, row, column);
        var display = CellDisplayFormatter.FormatDisplay(cell?.Value ?? CellValue.Blank, style, _book.DateSystem);
        if (display.Color is not null) style = style with { FontColor = display.Color };
        if (style.Horizontal == 0 && cell?.Value.Kind is CellValueKind.Number) style = style with { Horizontal = 3 };
        if (_sheet.Printing.Setup?.BlackAndWhite == true) style = style with { FillColor = "FFFFFFFF", FontColor = "FF000000", LeftBorderColor = "FF000000", RightBorderColor = "FF000000", TopBorderColor = "FF000000", BottomBorderColor = "FF000000" };
        dc.DrawRectangle(FastGridArtwork.Brush(style.FillColor, Brushes.White), _sheet.Printing.Options?.GridLines == true ? new Pen(Brushes.LightGray, .5) : null, rect);
        void Border(int kind, string? color, Point from, Point to)
        {
            if (kind == 0) return;
            var pen = new Pen(FastGridArtwork.Brush(color, Brushes.Black), kind is 2 or 5 ? 2 : 1);
            if (kind == 3) pen.DashStyle = DashStyles.Dash; if (kind == 4) pen.DashStyle = DashStyles.Dot;
            dc.DrawLine(pen, from, to);
            if (kind == 6) { var offset = from.X == to.X ? new Vector(2, 0) : new Vector(0, 2); dc.DrawLine(pen, from + offset, to + offset); }
        }
        Border(style.LeftBorder, style.LeftBorderColor, rect.TopLeft, rect.BottomLeft); Border(style.RightBorder, style.RightBorderColor, rect.TopRight, rect.BottomRight);
        Border(style.TopBorder, style.TopBorderColor, rect.TopLeft, rect.TopRight); Border(style.BottomBorder, style.BottomBorderColor, rect.BottomLeft, rect.BottomRight);
        var text = display.Text;
        if (cell?.Value.Kind == CellValueKind.Error) text = _sheet.Printing.Setup?.Errors switch { "blank" => "", "dash" => "--", "NA" => "#N/A", _ => text };
        if (display is { FillOffset: >= 0, FillCharacter: { } fillCharacter })
        {
            var unit = FastGridArtwork.CreateText(fillCharacter.ToString(), style, 1, null, 14).WidthIncludingTrailingWhitespace;
            var occupied = FastGridArtwork.CreateText(text, style, 1, null, 14).WidthIncludingTrailingWhitespace;
            var count = unit > 0 ? Math.Clamp((int)Math.Floor((rect.Width - 4 - occupied) / unit), 0, 4096) : 0;
            text = text.Insert(display.FillOffset, new string(fillCharacter, count));
        }
        Text(dc, text, rect, style, cell?.Content.RichText);
    }
    private static void Text(DrawingContext dc, string value, Rect rect, StyleDefinition style, RichCellText? rich = null)
    {
        if (rect.Width <= 4 || rect.Height <= 0) return;
        var text = FastGridArtwork.CreateText(value, style, 1, rich, 14); var indent = FastGridArtwork.IndentWidth(style, 1, 14); var width = Math.Max(1, rect.Width - 4 - indent);
        if (rich is null) FastGridArtwork.FitText(text, style, width, 14);
        text.MaxTextWidth = style.TextRotation == 0 ? width : Math.Max(width, text.WidthIncludingTrailingWhitespace); text.MaxLineCount = style.WrapText ? int.MaxValue : 1;
        text.TextAlignment = style.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left };
        var y = style.Vertical switch { 0 => rect.Top + 1, 1 => rect.Top + (rect.Height - text.Height) / 2, _ => rect.Bottom - text.Height - 1 };
        dc.PushClip(new RectangleGeometry(rect));
        if (style.TextRotation is > 0 and <= 180)
        {
            text.TextAlignment = TextAlignment.Left;
            var rotate = new RotateTransform(style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90);
            var bounds = rotate.TransformBounds(new Rect(0, 0, text.WidthIncludingTrailingWhitespace, text.Height));
            var x = style.Horizontal switch { 2 => (width - bounds.Width) / 2, 3 => width - bounds.Width, _ => 0 };
            var top = style.Vertical switch { 0 => 0, 1 => (rect.Height - bounds.Height) / 2, _ => rect.Height - bounds.Height };
            dc.PushTransform(new TranslateTransform(rect.Left + 2 + indent + x - bounds.Left, rect.Top + top - bounds.Top));
            dc.PushTransform(rotate); dc.DrawText(text, new(0, 0)); dc.Pop(); dc.Pop();
        }
        else dc.DrawText(text, new(rect.Left + 2 + (style.Horizontal == 3 ? 0 : indent), y));
        dc.Pop();
    }
    private void Header(DrawingContext dc, string? source, double y, int page)
    {
        var text = PrintingLayout.Header(source, _sheet.Name, _file, page, PageCount, _timestamp);
        var align = _sheet.Printing.HeaderFooter?.AlignWithMargins != false;
        var left = align ? Layout.Margins.Left * 96 : 24; var right = align ? Layout.Margins.Right * 96 : 24;
        var scale = _sheet.Printing.HeaderFooter?.ScaleWithDoc != false ? Layout.Scale : 1;
        var width = (PageSize.Width - left - right) / scale;
        dc.PushTransform(new TranslateTransform(left, y)); dc.PushTransform(new ScaleTransform(scale, scale));
        var rect = new Rect(0, 0, width, 20);
        Text(dc, text.Left, rect, new() { Horizontal = 1, FontSize = 10 }); Text(dc, text.Center, rect, new() { Horizontal = 2, FontSize = 10 }); Text(dc, text.Right, rect, new() { Horizontal = 3, FontSize = 10 });
        dc.Pop(); dc.Pop();
    }
}
