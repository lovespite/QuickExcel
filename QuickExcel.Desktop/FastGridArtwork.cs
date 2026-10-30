using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Core;
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

    internal void Refresh()
    {
        _sheet = surface.Sheet;
        _hasStyles = surface.Snapshot.Styles.Any(style => style != DefaultStyle);
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
        void Paint(CellRectangle range, Action<Rect> paint)
        {
            foreach (var (bounds, clip) in grid.GetModelRangeFragments(new(range.Top, range.Left, range.Bottom, range.Right)))
            {
                drawing.PushClip(new RectangleGeometry(clip)); paint(bounds); drawing.Pop();
            }
        }
        var columns = _hasStyles ? grid.GetVisibleModelColumns() : [];
        foreach (var row in _hasStyles ? grid.GetVisibleModelRows() : [])
        foreach (var column in columns)
        {
            var style = surface.Snapshot.Styles[_sheet.GetStyleId(row, column)];
            if (style == DefaultStyle) continue;
            Paint(new(row, column, row, column), rect => PaintCell(drawing, rect, style, surface.DisplayText(row, column), _sheet.ShowGridLines));
        }
        foreach (var merge in _merges)
            Paint(merge, rect => PaintCell(drawing, rect, surface.Snapshot.Styles[_sheet.GetStyleId(merge.Top, merge.Left)],
                surface.DisplayText(merge.Top, merge.Left), _sheet.ShowGridLines));
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

    private void PaintCell(DrawingContext drawing, Rect rect, StyleDefinition style, string value, bool gridLines)
    {
        drawing.DrawRectangle(Brush(style.FillColor, Brushes.White), null, rect);
        drawing.PushClip(new RectangleGeometry(rect));
        var text = CreateText(value, style, VisualTreeHelper.GetDpi(this).PixelsPerDip);
        // Keep the natural line box and let the cell clip trim glyphs. The same
        // font/layout settings are used to measure automatic row heights.
        text.MaxTextWidth = Math.Max(1, rect.Width - 4);
        text.MaxLineCount = style.WrapText ? int.MaxValue : 1;
        text.Trimming = TextTrimming.CharacterEllipsis;
        text.TextAlignment = style.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left };
        var y = style.Vertical switch { 0 => 1, 1 => (rect.Height - text.Height) / 2, _ => rect.Height - text.Height - 1 };
        drawing.DrawText(text, new(rect.Left + 2, rect.Top + y));
        if (gridLines) drawing.DrawRectangle(null, new Pen(Brushes.Gainsboro, 1), rect);
        void Border(int width, string? color, Point start, Point end)
        { if (width > 0) drawing.DrawLine(new Pen(Brush(color, Brushes.Gray), 1), start, end); }
        Border(style.LeftBorder, style.LeftBorderColor, rect.TopLeft, rect.BottomLeft);
        Border(style.TopBorder, style.TopBorderColor, rect.TopLeft, rect.TopRight);
        Border(style.RightBorder, style.RightBorderColor, rect.TopRight, rect.BottomRight);
        Border(style.BottomBorder, style.BottomBorderColor, rect.BottomLeft, rect.BottomRight);
        drawing.Pop();
    }
    internal static FormattedText CreateText(string value, StyleDefinition style, double pixelsPerDip) =>
        new(value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(Font(style.FontName), style.Italic ? FontStyles.Italic : FontStyles.Normal,
                style.Bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal),
            (style.FontSize ?? 11) * 96 / 72, Brush(style.FontColor, Brushes.Black), pixelsPerDip);

    internal static FontFamily Font(string? name)
    {
        try { return new(string.IsNullOrWhiteSpace(name) ? "Microsoft YaHei UI" : name); }
        catch (ArgumentException) { return new("Microsoft YaHei UI"); }
    }
    internal static Brush Brush(string? color, Brush fallback)
    {
        try { return string.IsNullOrEmpty(color) ? fallback : (Brush)new BrushConverter().ConvertFromString("#" + color.TrimStart('#'))!; }
        catch (Exception ex) when (ex is FormatException or NotSupportedException) { return fallback; }
    }
}
