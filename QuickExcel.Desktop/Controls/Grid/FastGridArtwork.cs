using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Core;
using QuickExcel.Calculation;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>View-only formatting and artwork, below the grid's selection layer.</summary>
internal sealed class FastGridArtwork(FastGridSurface surface) : FrameworkElement
{
    private readonly Dictionary<SheetImage, BitmapSource?> _images = [];
    private readonly Dictionary<SheetImage, (CellRectangle Range, DrawingBounds Bounds)> _imageGeometry = [];
    private CellRectangle[] _merges = [];
    private WorksheetSnapshot? _sheet;
    private bool _hasStyles;
    private static readonly StyleDefinition DefaultStyle = new();
    private readonly record struct CellDrawingKey(CellRectangle Range, double Width, double Height, double ScaleX, double ScaleY);
    private readonly Dictionary<CellDrawingKey, DrawingGroup> _cellDrawings = [];
    internal long CellDrawingCount { get; private set; }
    internal int CachedCellDrawingCount => _cellDrawings.Count;

    internal void Refresh()
    {
        // Values, styles, merges and gridlines belong to this immutable snapshot.
        // Geometry/DPI are part of each key; scrolling only translates the drawing.
        _cellDrawings.Clear();
        _sheet = surface.Sheet;
        _imageGeometry.Clear();
        foreach (var picture in _sheet.Images)
        {
            var bounds = picture.Anchor.Bounds(_sheet);
            if (bounds.Width <= 0 || bounds.Height <= 0) continue;
            var anchor = DrawingGeometry.FromBounds(_sheet, bounds, DrawingAnchorKind.TwoCell);
            _imageGeometry[picture] = (new(anchor.Row, anchor.Column, Math.Min(WorkbookLimits.MaxRows - 1, anchor.EndRow), Math.Min(WorkbookLimits.MaxColumns - 1, anchor.EndColumn)), bounds);
        }
        _hasStyles = surface.Snapshot.Styles.Any(style => style != DefaultStyle) || _sheet.Cells.Values.Any(c => c.Content.RichText is not null) || !_sheet.Annotations.IsEmpty || !_sheet.Data.Tables.IsEmpty || !_sheet.Data.ConditionalFormats.IsEmpty;
        _merges = _sheet.MergedCells.Select(CellRectangle.Parse).ToArray();
        foreach (var image in _images.Keys.Where(i => !_sheet.Images.Contains(i)).ToArray()) _images.Remove(image);
        foreach (var image in _sheet.Images.Where(i => !_images.ContainsKey(i)))
        {
            try
            {
                using var stream = new MemoryStream(image.Data.ToArray(), false);
                var header = BitmapDecoder.Create(stream, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None).Frames[0];
                if ((long)header.PixelWidth * header.PixelHeight > 16_000_000) { _images[image] = null; continue; }
                stream.Position = 0;
                var bitmap = new BitmapImage(); bitmap.BeginInit(); bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.DecodePixelWidth = Math.Min(2048, header.PixelWidth); bitmap.StreamSource = stream; bitmap.EndInit();
                bitmap.Freeze(); _images[image] = bitmap;
            }
            catch (Exception ex) when (ex is IOException or NotSupportedException or ArgumentException or System.Runtime.InteropServices.COMException)
            { _images[image] = null; }
        }
        InvalidateVisual();
    }

    protected override void OnRender(DrawingContext drawing)
    {
        if (_sheet is null) return;
        var grid = surface.Control;
        var used = new HashSet<CellDrawingKey>();
        void Paint(CellRectangle range, Action<Rect> paint)
        {
            foreach (var (bounds, clip) in grid.GetModelRangeFragments(new(range.Top, range.Left, range.Bottom, range.Right)))
            {
                drawing.PushClip(new RectangleGeometry(clip)); paint(bounds); drawing.Pop();
            }
        }
        void PaintCachedCell(CellRectangle range, StyleDefinition style)
        {
            Paint(range, rect =>
            {
                var key = new CellDrawingKey(range, rect.Width, rect.Height, grid.RenderScaleX, grid.RenderScaleY);
                used.Add(key);
                if (!_cellDrawings.TryGetValue(key, out var content))
                {
                    var display = CellDisplayFormatter.FormatDisplay(_sheet.GetCell(range.Top, range.Left)?.Value ?? CellValue.Blank,
                        style, surface.Snapshot.DateSystem);
                    content = new DrawingGroup();
                    using (var context = content.Open())
                        PaintCell(context, new Rect(0, 0, rect.Width, rect.Height),
                            display.Color is null ? style : style with { FontColor = display.Color }, display.Text, _sheet.ShowGridLines, display,
                            _sheet.GetCell(range.Top, range.Left)?.Content.RichText);
                    if (_sheet.Annotations.Notes.ContainsKey(new(range.Top, range.Left)))
                    {
                        using var context = content.Append();
                        var marker = new StreamGeometry(); using (var geometry = marker.Open())
                        { geometry.BeginFigure(new(rect.Width - 7, 1), true, true); geometry.LineTo(new(rect.Width - 1, 1), true, false); geometry.LineTo(new(rect.Width - 1, 7), true, false); }
                        context.DrawGeometry(Brushes.Red, null, marker);
                    }
                    content.Freeze();
                    // Keep immutable vector content. RenderTargetBitmap.Render
                    // here synchronously renders each newly exposed cell on the
                    // UI thread and churns native surfaces during thumb dragging.
                    _cellDrawings.Add(key, content);
                    CellDrawingCount++;
                }
                drawing.PushTransform(new TranslateTransform(rect.Left, rect.Top));
                drawing.DrawDrawing(content);
                drawing.Pop();
            });
        }
        var columns = _hasStyles ? grid.GetVisibleModelColumns() : [];
        foreach (var row in _hasStyles ? grid.GetVisibleModelRows() : [])
        foreach (var column in columns)
        {
            var style = ConditionalPresentation.GetStyle(surface.Snapshot, _sheet, row, column);
            if (style == DefaultStyle && _sheet.GetCell(row, column)?.Content.RichText is null && !_sheet.Annotations.Notes.ContainsKey(new(row, column))) continue;
            PaintCachedCell(new(row, column, row, column), style);
        }
        foreach (var merge in _merges)
            PaintCachedCell(merge, ConditionalPresentation.GetStyle(surface.Snapshot, _sheet, merge.Top, merge.Left));
        surface.ReportFormattingDiagnostics();
        // Keep only visible cells/merge fragments; traversing a sheet cannot grow
        // the cache with every visited address or retain a previous viewport's DPI.
        foreach (var key in _cellDrawings.Keys.Where(key => !used.Contains(key)).ToArray()) _cellDrawings.Remove(key);
        foreach (var image in _sheet.Images)
        {
            if (!_imageGeometry.TryGetValue(image, out var geometry) || geometry.Range.Top >= surface.RowCount || geometry.Range.Left >= surface.ColumnCount) continue;
            Paint(geometry.Range with { Bottom = Math.Min(surface.RowCount - 1, geometry.Range.Bottom), Right = Math.Min(surface.ColumnCount - 1, geometry.Range.Right) }, rect =>
            {
                var bounds = geometry.Bounds;
                var actual = new Rect(rect.Left + bounds.X - DrawingGeometry.ColumnStart(_sheet, geometry.Range.Left), rect.Top + bounds.Y - DrawingGeometry.RowStart(_sheet, geometry.Range.Top), bounds.Width, bounds.Height);
                if (_images.GetValueOrDefault(image) is { } bitmap) drawing.DrawImage(bitmap, actual);
                else PaintCell(drawing, actual, DefaultStyle, "图片无法预览", true);
            });
        }
        foreach (var chart in _sheet.Data.Charts)
        {
            var bounds = chart.Anchor.Bounds(_sheet); if (bounds.Width <= 0 || bounds.Height <= 0) continue;
            var anchor = DrawingGeometry.FromBounds(_sheet, bounds, DrawingAnchorKind.TwoCell);
            var range = new CellRectangle(anchor.Row, anchor.Column, Math.Min(surface.RowCount - 1, anchor.EndRow), Math.Min(surface.ColumnCount - 1, anchor.EndColumn));
            if (range.Top > range.Bottom || range.Left > range.Right) continue;
            Paint(range, rect => ChartDrawing.Paint(drawing, new Rect(rect.Left + bounds.X - DrawingGeometry.ColumnStart(_sheet, range.Left), rect.Top + bounds.Y - DrawingGeometry.RowStart(_sheet, range.Top), bounds.Width, bounds.Height), surface.Snapshot, chart));
        }
    }
    internal static bool PaintsStyle(StyleDefinition style) => style != DefaultStyle;

    private void PaintCell(DrawingContext drawing, Rect rect, StyleDefinition style, string value, bool gridLines, CellDisplay? display = null, RichCellText? rich = null)
    {
        var grid = surface.Control;
        var sx = grid.RenderScaleX > 0 ? grid.RenderScaleX : 1;
        var sy = grid.RenderScaleY > 0 ? grid.RenderScaleY : 1;
        var px = 1 / sx;
        var py = 1 / sy;
        // The raster layer paints a cell as [left, right] with its separator on
        // the leading pixel and the fill one pixel inside. Reproduce that
        // geometry with whole device pixels, otherwise a vector fill/border
        // drifts by 1-2 px from the neighbouring raster cells.
        rect = new Rect(Snap(rect.X, sx), Snap(rect.Y, sy), Snap(rect.Width, sx), Snap(rect.Height, sy));
        var fill = new Rect(rect.X + px, rect.Y + py, Math.Max(0, rect.Width - px), Math.Max(0, rect.Height - py));
        drawing.DrawRectangle(Brush(style.FillColor, Brushes.White), null, fill);

        var padX = grid.CellPaddingHorizontal * px;
        var padY = grid.CellPaddingVertical * py;
        // The rasterizer's cell rect includes the trailing separator pixel, so
        // its content box is inset by the padding on the leading sides only.
        var content = new Rect(rect.X + padX, rect.Y + padY,
            Math.Max(0, rect.Width - padX - px), Math.Max(0, rect.Height - padY));
        var pixelsPerDip = sy;
        var defaultFontSize = grid.CellFontSize;
        var text = CreateText(value, style, pixelsPerDip, rich, defaultFontSize);
        if (display is { FillOffset: >= 0, FillCharacter: { } fillCharacter })
        {
            var unit = CreateText(fillCharacter.ToString(), style, pixelsPerDip, null, defaultFontSize).WidthIncludingTrailingWhitespace;
            var count = unit > 0 ? Math.Clamp((int)Math.Floor((content.Width - text.WidthIncludingTrailingWhitespace) / unit), 0, 4096) : 0;
            text = CreateText(value.Insert(display.FillOffset, new string(fillCharacter, count)), style, pixelsPerDip, null, defaultFontSize);
        }
        // Keep the natural line box and let the cell clip trim glyphs. The same
        // font/layout settings are used to measure automatic row heights.
        var indent = IndentWidth(style, pixelsPerDip, defaultFontSize);
        var available = Math.Max(px, content.Width - indent);
        if (rich is null) FitText(text, style, available, defaultFontSize);
        text.MaxTextWidth = style.TextRotation == 0 ? available : Math.Max(available, text.WidthIncludingTrailingWhitespace);
        text.MaxLineCount = style.WrapText ? int.MaxValue : 1;
        text.Trimming = style.ShrinkToFit && !style.WrapText ? TextTrimming.None : TextTrimming.CharacterEllipsis;
        text.TextAlignment = style.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left };
        // Vertical placement must use the rasterizer's metrics and rounding:
        // it centres ceil(line height in device pixels) inside the padded
        // content box, so a plain vector centre sinks the text by a pixel.
        var textHeightPx = Math.Ceiling(text.Height * sy);
        var contentTopPx = Math.Round(rect.Y * sy) + grid.CellPaddingVertical;
        var contentHeightPx = Math.Round(rect.Height * sy) - grid.CellPaddingVertical;
        var topPx = style.Vertical switch
        {
            0 => contentTopPx,
            1 => contentTopPx + (int)Math.Round(contentHeightPx / 2.0 - textHeightPx / 2.0),
            _ => contentTopPx + contentHeightPx - (int)textHeightPx,
        };
        drawing.PushClip(new RectangleGeometry(content));
        if (style.TextRotation is > 0 and <= 180)
        {
            text.TextAlignment = TextAlignment.Left;
            var rotate = new RotateTransform(style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90);
            var bounds = rotate.TransformBounds(new Rect(0, 0, text.WidthIncludingTrailingWhitespace, text.Height));
            var x = style.Horizontal switch { 2 => (available - bounds.Width) / 2, 3 => available - bounds.Width, _ => 0 };
            var y = style.Vertical switch { 0 => 0, 1 => (content.Height - bounds.Height) / 2, _ => content.Height - bounds.Height };
            drawing.PushTransform(new TranslateTransform(content.Left + indent + x - bounds.Left, content.Top + y - bounds.Top));
            drawing.PushTransform(rotate); drawing.DrawText(text, new(0, 0)); drawing.Pop(); drawing.Pop();
        }
        else
        {
            drawing.DrawText(text, new(content.Left + (style.Horizontal == 3 ? 0 : indent), topPx / sy));
        }
        drawing.Pop();
        // Gridlines and cell borders share the rasterizer's edge convention:
        // a one-device-pixel band starting at each side (the trailing band
        // belongs to the following cell's leading separator). The rasterizer
        // draws this band even when gridlines are disabled, using white.
        void Band(Brush brush, double x, double y, double width, double height) =>
            drawing.DrawRectangle(brush, null, new Rect(x, y, Math.Max(px, width), Math.Max(py, height)));
        var separator = new SolidColorBrush(gridLines ? grid.GridLineColor : Colors.White); separator.Freeze();
        Band(separator, rect.X, rect.Y, rect.Width + px, py);
        Band(separator, rect.X, rect.Y, px, rect.Height + py);
        Band(separator, rect.X, rect.Bottom, rect.Width + px, py);
        Band(separator, rect.Right, rect.Y, px, rect.Height + py);
        void Border(int width, string? color, bool horizontal, bool trailing)
        {
            if (width <= 0) return;
            // Border styles are drawn as one device pixel; the style name only
            // decides the colour here, matching the previous projection.
            var brush = Brush(color, Brushes.Gray);
            if (horizontal) Band(brush, rect.X, trailing ? rect.Bottom : rect.Y, rect.Width + px, py);
            else Band(brush, trailing ? rect.Right : rect.X, rect.Y, px, rect.Height + py);
        }
        Border(style.LeftBorder, style.LeftBorderColor, false, false);
        Border(style.TopBorder, style.TopBorderColor, true, false);
        Border(style.RightBorder, style.RightBorderColor, false, true);
        Border(style.BottomBorder, style.BottomBorderColor, true, true);
    }

    /// <summary>Rounds a DIP coordinate or length to the device pixel grid.</summary>
    private static double Snap(double value, double scale) => Math.Round(value * scale) / scale;
    internal static FormattedText CreateText(string value, StyleDefinition style, double pixelsPerDip, RichCellText? rich = null, double defaultFontSize = 0)
    {
        var text = new FormattedText(value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(Font(style.FontName, style.EastAsianFontName), style.Italic ? FontStyles.Italic : FontStyles.Normal,
                style.Bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal),
            FontPixels(style, defaultFontSize), Brush(style.FontColor, Brushes.Black), pixelsPerDip);
        text.SetTextDecorations(Decorations(style.Underline, style.Strike));
        if (rich is null || rich.Text != value) return text;
        var offset = 0;
        foreach (var run in rich.Runs)
        {
            var length = run.Text.Length;
            if (length > 0 && run.Font is { } font)
            {
                if (font.ResolvedName is not null || font.EastAsianName is not null) text.SetFontFamily(Font(font.ResolvedName ?? style.FontName, font.EastAsianName ?? style.EastAsianFontName), offset, length);
                if (font.Size is { } size) text.SetFontSize(size * 96 / 72, offset, length);
                if (font.Color is { } color) text.SetForegroundBrush(Brush(color, Brushes.Black), offset, length);
                else if (font.IndexedColor == 81) text.SetForegroundBrush(SystemColors.InfoTextBrush, offset, length);
                if (font.Bold is { } bold) text.SetFontWeight(bold ? FontWeights.Bold : FontWeights.Normal, offset, length);
                if (font.Italic is { } italic) text.SetFontStyle(italic ? FontStyles.Italic : FontStyles.Normal, offset, length);
                text.SetTextDecorations(Decorations(font.Underline ?? style.Underline, font.Strike ?? style.Strike), offset, length);
            }
            offset += length;
        }
        return text;
    }
    private static TextDecorationCollection Decorations(string? underline, bool strike)
    {
        var result = new TextDecorationCollection();
        if (underline is not (null or "none")) result.Add(TextDecorations.Underline[0]);
        if (underline is "double" or "doubleAccounting") result.Add(new TextDecoration(TextDecorationLocation.Underline, null, 2, TextDecorationUnit.Pixel, TextDecorationUnit.FontRecommended));
        if (strike) result.Add(TextDecorations.Strikethrough[0]);
        return result;
    }

    internal static double IndentWidth(StyleDefinition style, double pixelsPerDip, double defaultFontSize = 0) => style.Indent == 0 ? 0 :
        CreateText(" ", style, pixelsPerDip, null, defaultFontSize).WidthIncludingTrailingWhitespace * 3 * style.Indent;

    /// <summary>Points stored by the workbook converted to DIP; unformatted cells use the grid font.</summary>
    internal static double FontPixels(StyleDefinition style, double defaultFontSize = 0) =>
        style.FontSize is { } points ? points * 96 / 72 : defaultFontSize > 0 ? defaultFontSize : 11.0 * 96 / 72;

    internal static void FitText(FormattedText text, StyleDefinition style, double available, double defaultFontSize = 0)
    {
        if (!style.ShrinkToFit || style.WrapText) return;
        var size = FontPixels(style, defaultFontSize);
        // Hinting is not exactly proportional to size. Re-measure after each bounded reduction.
        for (var i = 0; i < 16 && size > 1 && text.WidthIncludingTrailingWhitespace > available; i++)
        {
            size = Math.Max(1, size * Math.Min(.97, Math.Max(1, available - 2) / text.WidthIncludingTrailingWhitespace));
            text.SetFontSize(size);
        }
    }

    internal static FontFamily Font(string? name, string? fallback = null)
    {
        try { return new(string.IsNullOrWhiteSpace(name) ? "Microsoft YaHei UI" : string.IsNullOrEmpty(fallback) ? name : name + ", " + fallback); }
        catch (ArgumentException) { return new("Microsoft YaHei UI"); }
    }
    internal static Brush Brush(string? color, Brush fallback)
    {
        try { return string.IsNullOrEmpty(color) ? fallback : (Brush)new BrushConverter().ConvertFromString("#" + color.TrimStart('#'))!; }
        catch (Exception ex) when (ex is FormatException or NotSupportedException) { return fallback; }
    }
}
