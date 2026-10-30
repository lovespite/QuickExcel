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
    private CellRectangle[] _merges = [];
    private WorksheetSnapshot? _sheet;
    private bool _hasStyles;
    private static readonly StyleDefinition DefaultStyle = new();
    private readonly record struct CellDrawingKey(CellRectangle Range, double Width, double Height, double DpiX, double DpiY);
    private readonly Dictionary<CellDrawingKey, DrawingGroup> _cellDrawings = [];
    internal long CellDrawingCount { get; private set; }
    internal int CachedCellDrawingCount => _cellDrawings.Count;

    internal void Refresh()
    {
        // Values, styles, merges and gridlines belong to this immutable snapshot.
        // Geometry/DPI are part of each key; scrolling only translates the drawing.
        _cellDrawings.Clear();
        _sheet = surface.Sheet;
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
        var dpi = VisualTreeHelper.GetDpi(this);
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
                var key = new CellDrawingKey(range, rect.Width, rect.Height, dpi.DpiScaleX, dpi.DpiScaleY);
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
            if (image.RowIndex >= surface.RowCount || image.ColumnIndex >= surface.ColumnCount) continue;
            Paint(new(image.RowIndex, image.ColumnIndex, Math.Min(surface.RowCount, image.RowIndex + image.RowSpan) - 1,
                Math.Min(surface.ColumnCount, image.ColumnIndex + image.ColumnSpan) - 1), rect =>
            {
                if (_images.GetValueOrDefault(image) is { } bitmap) drawing.DrawImage(bitmap, rect);
                else PaintCell(drawing, rect, DefaultStyle, "图片无法预览", true);
            });
        }
    }
    internal static bool PaintsStyle(StyleDefinition style) => style != DefaultStyle;

    private void PaintCell(DrawingContext drawing, Rect rect, StyleDefinition style, string value, bool gridLines, CellDisplay? display = null, RichCellText? rich = null)
    {
        drawing.DrawRectangle(Brush(style.FillColor, Brushes.White), null, rect);
        drawing.PushClip(new RectangleGeometry(rect));
        var pixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var text = CreateText(value, style, pixelsPerDip, rich);
        if (display is { FillOffset: >= 0, FillCharacter: { } fill })
        {
            var unit = CreateText(fill.ToString(), style, pixelsPerDip).WidthIncludingTrailingWhitespace;
            var count = unit > 0 ? Math.Clamp((int)Math.Floor((rect.Width - 4 - text.WidthIncludingTrailingWhitespace) / unit), 0, 4096) : 0;
            text = CreateText(value.Insert(display.FillOffset, new string(fill, count)), style, pixelsPerDip);
        }
        // Keep the natural line box and let the cell clip trim glyphs. The same
        // font/layout settings are used to measure automatic row heights.
        var indent = IndentWidth(style, pixelsPerDip);
        var available = Math.Max(1, rect.Width - 4 - indent);
        if (rich is null) FitText(text, style, available);
        text.MaxTextWidth = style.TextRotation == 0 ? available : Math.Max(available, text.WidthIncludingTrailingWhitespace);
        text.MaxLineCount = style.WrapText ? int.MaxValue : 1;
        text.Trimming = style.ShrinkToFit && !style.WrapText ? TextTrimming.None : TextTrimming.CharacterEllipsis;
        text.TextAlignment = style.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left };
        if (style.TextRotation is > 0 and <= 180)
        {
            text.TextAlignment = TextAlignment.Left;
            var rotate = new RotateTransform(style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90);
            var bounds = rotate.TransformBounds(new Rect(0, 0, text.WidthIncludingTrailingWhitespace, text.Height));
            var x = style.Horizontal switch { 2 => (available - bounds.Width) / 2, 3 => available - bounds.Width, _ => 0 };
            var y = style.Vertical switch { 0 => 1, 1 => (rect.Height - bounds.Height) / 2, _ => rect.Height - bounds.Height - 1 };
            drawing.PushTransform(new TranslateTransform(rect.Left + 2 + indent + x - bounds.Left, rect.Top + y - bounds.Top));
            drawing.PushTransform(rotate); drawing.DrawText(text, new(0, 0)); drawing.Pop(); drawing.Pop();
        }
        else
        {
            var y = style.Vertical switch { 0 => 1, 1 => (rect.Height - text.Height) / 2, _ => rect.Height - text.Height - 1 };
            drawing.DrawText(text, new(rect.Left + 2 + (style.Horizontal == 3 ? 0 : indent), rect.Top + y));
        }
        if (gridLines) drawing.DrawRectangle(null, new Pen(Brushes.Gainsboro, 1), rect);
        void Border(int width, string? color, Point start, Point end)
        { if (width > 0) drawing.DrawLine(new Pen(Brush(color, Brushes.Gray), 1), start, end); }
        Border(style.LeftBorder, style.LeftBorderColor, rect.TopLeft, rect.BottomLeft);
        Border(style.TopBorder, style.TopBorderColor, rect.TopLeft, rect.TopRight);
        Border(style.RightBorder, style.RightBorderColor, rect.TopRight, rect.BottomRight);
        Border(style.BottomBorder, style.BottomBorderColor, rect.BottomLeft, rect.BottomRight);
        drawing.Pop();
    }
    internal static FormattedText CreateText(string value, StyleDefinition style, double pixelsPerDip, RichCellText? rich = null)
    {
        var text = new FormattedText(value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(Font(style.FontName, style.EastAsianFontName), style.Italic ? FontStyles.Italic : FontStyles.Normal,
                style.Bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal),
            (style.FontSize ?? 11) * 96 / 72, Brush(style.FontColor, Brushes.Black), pixelsPerDip);
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

    internal static double IndentWidth(StyleDefinition style, double pixelsPerDip) => style.Indent == 0 ? 0 :
        CreateText(" ", style, pixelsPerDip).WidthIncludingTrailingWhitespace * 3 * style.Indent;

    internal static void FitText(FormattedText text, StyleDefinition style, double available)
    {
        if (!style.ShrinkToFit || style.WrapText) return;
        var size = (style.FontSize ?? 11) * 96 / 72;
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
