using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Domain;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

/// <summary>Draws model rectangles over realized grid rows; never writes to the workbook.</summary>
internal sealed class SheetArtworkAdorner(DataGrid grid, WpfGridSurface surface, bool frozen) : Adorner(grid)
{
    private readonly Dictionary<SheetImage, BitmapSource?> _images = [];
    private double[] _rowOffsets = [];
    private WorksheetSnapshot? _sheet;
    private CellRectangle[] _merges = [];
    private int _layoutHash;
    private int _previewRow = -1;
    private double _previewDelta;

    public void PreviewRowHeight(int row, double delta)
    {
        _previewRow = row; _previewDelta = delta;
        if (_rowOffsets.Length > 0) InvalidateVisual();
    }

    private double RowOffset(int row) => _rowOffsets[row] + (_previewRow >= 0 && row > _previewRow ? _previewDelta : 0);

    // Adorners may render before rows/columns are realized after a page or dock change.
    // Invalidate only when their geometry changes, never on every LayoutUpdated event.
    public void RefreshLayout()
    {
        if (_rowOffsets.Length == 0) return;
        var hash = new HashCode();
        hash.Add(grid.IsVisible); hash.Add(grid.RenderSize); hash.Add(RenderSize);
        foreach (var column in grid.Columns) { hash.Add(column.ActualWidth); hash.Add(column.Visibility); }
        foreach (var row in Children<DataGridRow>(grid))
        {
            hash.Add((row.Item as SheetRow)?.Index); hash.Add(row.ActualHeight);
            hash.Add(row.TranslatePoint(new Point(), grid));
        }
        var next = hash.ToHashCode();
        if (next == _layoutHash) return;
        _layoutHash = next; InvalidateVisual();
    }

    public void Refresh()
    {
        IsHitTestVisible = false;
        _previewRow = -1; _previewDelta = 0;
        _sheet = surface.Sheet;
        _merges = _sheet.MergedCells.Select(CellRectangle.Parse).ToArray();
        foreach (var key in _images.Keys.Where(image => !_sheet.Images.Contains(image)).ToArray()) _images.Remove(key);
        foreach (var image in _sheet.Images.Where(image => !_images.ContainsKey(image)))
        {
            try
            {
                using var stream = new MemoryStream(image.Data.ToArray(), false);
                var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None);
                var header = decoder.Frames[0];
                if ((long)header.PixelWidth * header.PixelHeight > 16_000_000) { _images[image] = null; continue; }
                stream.Position = 0;
                var bitmap = new BitmapImage(); bitmap.BeginInit(); bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.DecodePixelWidth = Math.Min(2048, header.PixelWidth); bitmap.StreamSource = stream; bitmap.EndInit();
                bitmap.Freeze(); _images[image] = bitmap;
            }
            catch (Exception ex) when (ex is IOException or NotSupportedException or ArgumentException or System.Runtime.InteropServices.COMException)
            { _images[image] = null; }
        }
        if (_merges.Length == 0 && _images.Count == 0) { _rowOffsets = []; InvalidateVisual(); return; }
        _rowOffsets = new double[surface.Rows.Count + 1];
        for (var row = 0; row < surface.Rows.Count; row++)
        {
            var hidden = _sheet.Rows.GetValueOrDefault(row)?.Hidden == true || surface.IsFiltered(row);
            _rowOffsets[row + 1] = _rowOffsets[row] + (hidden ? 0 : Math.Clamp(_sheet.Rows.GetValueOrDefault(row)?.Height ?? _sheet.DefaultRowHeight ?? 31, 0, 400));
        }
        InvalidateVisual();
    }

    protected override void OnRender(DrawingContext drawing)
    {
        if (_sheet is null || _rowOffsets.Length == 0 || !grid.IsVisible) return;
        var rows = Children<DataGridRow>(grid).Where(r => r.Item is SheetRow && r.ActualHeight > 0).ToArray();
        if (rows.Length == 0) return;
        var first = rows.OrderBy(r => ((SheetRow)r.Item).Index).First();
        var firstIndex = ((SheetRow)first.Item).Index;
        var originY = first.TranslatePoint(new Point(0, 0), grid).Y - RowOffset(firstIndex);
        var scroll = Children<ScrollViewer>(grid).FirstOrDefault();
        var top = frozen ? 0 : grid.ColumnHeaderHeight;
        var viewport = new Rect(grid.RowHeaderWidth, top, Math.Max(0, grid.ActualWidth - grid.RowHeaderWidth - SystemParameters.VerticalScrollBarWidth), Math.Max(0, grid.ActualHeight - top - (frozen ? 0 : SystemParameters.HorizontalScrollBarHeight)));
        var widths = new double[grid.Columns.Count + 1];
        for (var i = 0; i < grid.Columns.Count; i++) widths[i + 1] = widths[i] + (grid.Columns[i].Visibility == Visibility.Visible ? grid.Columns[i].ActualWidth : 0);
        var pinned = Math.Min(grid.FrozenColumnCount, grid.Columns.Count);
        var split = grid.RowHeaderWidth + widths[pinned];
        var offset = scroll?.HorizontalOffset ?? 0;
        var lastFrozen = frozen ? surface.View.FrozenRows : int.MaxValue;
        void Draw(CellRectangle range, Action<Rect> paint)
        {
            if (range.Top >= _rowOffsets.Length - 1 || range.Left >= grid.Columns.Count || range.Top >= lastFrozen) return;
            // Keep the full object geometry; the frozen viewport crops it instead of rescaling it.
            var bottom = Math.Min(range.Bottom + 1, _rowOffsets.Length - 1);
            var right = Math.Min(range.Right + 1, grid.Columns.Count);
            var y = originY + RowOffset(range.Top); var height = RowOffset(bottom) - RowOffset(range.Top);
            if (height <= 0 || widths[right] <= widths[range.Left]) return;
            foreach (var fixedPane in new[] { true, false })
            {
                if (fixedPane && range.Left >= pinned || !fixedPane && right <= pinned) continue;
                var clip = fixedPane ? new Rect(viewport.Left, viewport.Top, Math.Max(0, Math.Min(split, viewport.Right) - viewport.Left), viewport.Height)
                    : new Rect(Math.Max(viewport.Left, Math.Min(split, viewport.Right)), viewport.Top, Math.Max(0, viewport.Right - Math.Max(split, viewport.Left)), viewport.Height);
                var rectangle = new Rect(grid.RowHeaderWidth + widths[range.Left] - (fixedPane ? 0 : offset), y, widths[right] - widths[range.Left], height);
                if (!clip.IntersectsWith(rectangle)) continue;
                drawing.PushClip(new RectangleGeometry(clip)); paint(rectangle); drawing.Pop();
            }
        }
        foreach (var merge in _merges)
        {
            Draw(merge, rect =>
            {
                var cell = _sheet.GetCell(merge.Top, merge.Left);
                var style = surface.Snapshot.Styles[_sheet.GetStyleId(merge.Top, merge.Left)];
                var current = grid.CurrentItem is SheetRow row && row.Index == merge.Top && grid.CurrentColumn?.DisplayIndex == merge.Left;
                var editing = current && grid.CurrentColumn?.GetCellContent(grid.CurrentItem) is TextBox editor && !(bool)editor.GetValue(GridInputColumn.IsPreparedProperty);
                var stroke = current ? new SolidColorBrush(Color.FromRgb(24, 117, 83)) : new SolidColorBrush(Color.FromRgb(221, 229, 225));
                if (!editing)
                {
                    drawing.DrawRectangle(ColorBrush(style.FillColor, Brushes.White), null, rect);
                    drawing.PushClip(new RectangleGeometry(rect));
                    var text = new FormattedText(CellDisplayFormatter.Format(cell?.Value ?? CellValue.Blank, style, surface.Snapshot.DateSystem), CultureInfo.CurrentCulture,
                        FlowDirection.LeftToRight, new Typeface(new FontFamily("Microsoft YaHei UI"), style.Italic ? FontStyles.Italic : FontStyles.Normal,
                        style.Bold ? FontWeights.SemiBold : FontWeights.Normal, FontStretches.Normal), 13, ColorBrush(style.FontColor, Brushes.Black), VisualTreeHelper.GetDpi(this).PixelsPerDip)
                    { MaxTextWidth = Math.Max(1, rect.Width - 16), MaxTextHeight = Math.Max(1, rect.Height - 4), Trimming = TextTrimming.CharacterEllipsis,
                        TextAlignment = style.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left } };
                    drawing.DrawText(text, new Point(rect.X + 8, rect.Y + 4)); drawing.Pop();
                }
                drawing.DrawRectangle(null, new Pen(stroke, current ? 1.5 : 1), rect);
            });
        }
        foreach (var image in _sheet.Images)
        {
            Draw(new(image.RowIndex, image.ColumnIndex, image.RowIndex + image.RowSpan - 1, image.ColumnIndex + image.ColumnSpan - 1), rect =>
            {
                if (_images.GetValueOrDefault(image) is { } bitmap) drawing.DrawImage(bitmap, rect);
                else
                {
                    drawing.DrawRectangle(Brushes.WhiteSmoke, new Pen(Brushes.Gray, 1), rect);
                    drawing.PushClip(new RectangleGeometry(rect));
                    drawing.DrawText(new FormattedText("图片无法预览", CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Microsoft YaHei UI"), 12, Brushes.DimGray,
                        VisualTreeHelper.GetDpi(this).PixelsPerDip), rect.TopLeft + new Vector(4, 4)); drawing.Pop();
                }
            });
        }
    }
    private static Brush ColorBrush(string? color, Brush fallback)
    {
        if (string.IsNullOrEmpty(color)) return fallback;
        try { return (Brush)new BrushConverter().ConvertFromString("#" + color.TrimStart('#'))!; }
        catch (Exception ex) when (ex is FormatException or NotSupportedException) { return fallback; }
    }

    private static IEnumerable<T> Children<T>(DependencyObject parent) where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T match) yield return match;
            else foreach (var nested in Children<T>(child)) yield return nested;
        }
    }
}
