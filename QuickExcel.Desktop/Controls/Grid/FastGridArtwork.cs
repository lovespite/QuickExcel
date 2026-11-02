using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Core;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Desktop;

/// <summary>Immutable Skia pictures retain visible content; only a bounded viewport is rasterized.</summary>
internal sealed class FastGridArtwork : FrameworkElement, IDisposable
{
    private readonly FastGridSurface _surface;
    private bool _paintPending;
    internal bool IsRenderPending => _paintPending && IsVisible;
    private SkiaRenderer? _renderer;
    internal SkiaRenderer Renderer => _renderer ??= new();
    private readonly Dictionary<SheetImage, BitmapSource?> _images = [];
    private long _imageBytes;
    private CellRectangle[] _merges = [];
    private WorksheetSnapshot? _sheet;
    private readonly record struct CellDrawingKey(CellRectangle Range, double Width, double Height, double ScaleX, double ScaleY);
    private readonly Dictionary<CellDrawingKey, SKPicture> _cellDrawings = [];
    private readonly Dictionary<WorksheetChart, SKPicture> _charts = [];
    private readonly Dictionary<CellDrawingKey, BitmapSource> _cellRasters = [];
    private readonly Dictionary<(WorksheetChart Chart, double ScaleX, double ScaleY), BitmapSource> _chartRasters = [];
    private long _rasterBytes;
    private long _pictureBytes;
    private sealed record Pane(CellRectangle Anchor, Rect Bounds, Rect Clip, DrawingGroup Drawing);
    private Pane[]? _panes;
    private int[] _sceneRows = [], _sceneColumns = [];
    private double[] _sceneHeights = [], _sceneWidths = [];
    private (double Width, double Height, double ScaleX, double ScaleY) _sceneSize;
    private sealed class PaneRecording : IDisposable
    {
        internal readonly Dictionary<Rect, (Pane Pane, DrawingContext Context)> Items = [];
        public void Dispose() { foreach (var item in Items.Values) item.Context.Close(); }
    }
    internal long CellDrawingCount { get; private set; }
    internal int CachedCellDrawingCount => _cellDrawings.Count;
    internal long CachedNativeBytes => _pictureBytes + _imageBytes + _rasterBytes;
    internal double RecordMilliseconds { get; private set; }
    internal double RasterMilliseconds { get; private set; }
    internal double RenderMilliseconds { get; private set; }
    internal event Action<string, long>? RenderingTrace;
    internal FastGridArtwork(FastGridSurface surface)
    {
        _surface = surface;
        // Rasters already use the host's exact device scale. A second WPF bitmap
        // filtering pass is unnecessary and can blur the shaped glyph edges.
        RenderOptions.SetBitmapScalingMode(this, BitmapScalingMode.NearestNeighbor);
        Unloaded += (_, _) => Release();
        Loaded += (_, _) => InvalidateContent();
        SizeChanged += (_, _) => InvalidateContent();
    }
    private void ClearPictures()
    {
        _panes = null;
        foreach (var picture in _cellDrawings.Values) picture.Dispose();
        foreach (var picture in _charts.Values) picture.Dispose();
        _cellDrawings.Clear(); _charts.Clear(); _pictureBytes = 0;
        _cellRasters.Clear(); _chartRasters.Clear(); _rasterBytes = 0;
    }
    private void Release()
    {
        ClearPictures();
        _images.Clear(); _imageBytes = 0; _renderer?.Dispose(); _renderer = null;
    }
    public void Dispose() => Release();
    internal void Refresh()
    {
        // A snapshot owns all formatting and content; geometry is kept separately in the key.
        ClearPictures(); _sheet = _surface.Sheet;
        foreach (var key in _images.Keys.Where(key => !_sheet.Images.Contains(key)).ToArray())
        { _imageBytes -= Bytes(_images[key]); _images.Remove(key); }
        _merges = _sheet.MergedCells.Select(CellRectangle.Parse).ToArray();
        InvalidateContent();
    }
    internal void InvalidateContent() { _paintPending = true; InvalidateVisual(); }
    protected override void OnRender(DrawingContext drawing)
    {
        try { Paint(drawing); }
        finally { _paintPending = false; _surface.Control.NotifyContentRendered(); }
    }
    private void Paint(DrawingContext drawing)
    {
        if (_sheet is null || ActualWidth <= 0 || ActualHeight <= 0) return;
        var started = Stopwatch.GetTimestamp();
        RenderingTrace?.Invoke("artwork-start", started);
        var grid = _surface.Control;
        var sx = grid.RenderScaleX; var sy = grid.RenderScaleY;
        var rows = grid.GetVisibleModelRows().ToArray(); var columns = grid.GetVisibleModelColumns().ToArray();
        if (rows.Length == 0 || columns.Length == 0) return;
        var heights = rows.Select(row => grid.GetModelCellBounds(row, columns[0])?.Height ?? 0).ToArray();
        var widths = columns.Select(column => grid.GetModelCellBounds(rows[0], column)?.Width ?? 0).ToArray();
        var size = (ActualWidth, ActualHeight, sx, sy);
        var output = drawing;
        bool DrawPanes(Pane[] panes)
        {
            var offsets = new Vector[panes.Length];
            for (var i = 0; i < panes.Length; i++)
            {
                var pane = panes[i]; var anchor = pane.Anchor;
                var fragments = grid.GetModelRangeFragments(new(anchor.Top, anchor.Left, anchor.Bottom, anchor.Right));
                var current = fragments.FirstOrDefault(fragment => fragment.Clip == pane.Clip);
                if (current.Bounds.IsEmpty || current.Bounds.Width != pane.Bounds.Width || current.Bounds.Height != pane.Bounds.Height) return false;
                offsets[i] = current.Bounds.TopLeft - pane.Bounds.TopLeft;
            }
            for (var i = 0; i < panes.Length; i++)
            {
                var pane = panes[i]; var offset = offsets[i];
                output.PushClip(new RectangleGeometry(pane.Clip));
                output.PushTransform(new TranslateTransform(offset.X, offset.Y));
                output.DrawDrawing(pane.Drawing); output.Pop(); output.Pop();
            }
            return true;
        }
        // Retain one drawing per frozen/scrolling pane while the visible cell set is
        // unchanged. Fractional scrolling then translates existing immutable rasters
        // instead of recreating hundreds of WPF clip/image commands every frame.
        if (_panes is { } cached && _sceneSize == size && rows.SequenceEqual(_sceneRows) && columns.SequenceEqual(_sceneColumns) &&
            heights.SequenceEqual(_sceneHeights) && widths.SequenceEqual(_sceneWidths) && DrawPanes(cached))
        { RenderMilliseconds += Stopwatch.GetElapsedTime(started).TotalMilliseconds; RenderingTrace?.Invoke("artwork-cached-end", Stopwatch.GetTimestamp()); return; }
        _panes = null;
        using var recordingPanes = new PaneRecording();
        var reusable = true;
        var used = new HashSet<CellDrawingKey>(); var usedImages = new HashSet<SheetImage>(); var usedCharts = new HashSet<WorksheetChart>();
        void Paint(CellRectangle range, Action<Rect, Rect> paint)
        {
            foreach (var (bounds, clip) in grid.GetModelRangeFragments(new(range.Top, range.Left, range.Bottom, range.Right)))
            {
                if (!recordingPanes.Items.TryGetValue(clip, out var item))
                {
                    var group = new DrawingGroup();
                    item = (new(range, bounds, clip, group), group.Open()); recordingPanes.Items.Add(clip, item);
                }
                drawing = item.Context; paint(bounds, clip);
            }
        }
        void PaintCell(CellRectangle range)
        {
            Paint(range, (rect, clip) =>
            {
                var key = new CellDrawingKey(range, rect.Width, rect.Height, sx, sy); used.Add(key);
                if (!_cellDrawings.TryGetValue(key, out var content))
                {
                    var recording = Stopwatch.GetTimestamp();
                    var style = ConditionalPresentation.GetStyle(_surface.Snapshot, _sheet, range.Top, range.Left);
                    var cell = _sheet.GetCell(range.Top, range.Left);
                    _ = _surface.DisplayText(range.Top, range.Left); // Visible projection/read instrumentation.
                    var display = CellDisplayFormatter.FormatDisplay(cell?.Value ?? CellValue.Blank, style, _surface.Snapshot.DateSystem);
                    if (display.Color is not null) style = style with { FontColor = display.Color };
                    if (style.Horizontal == 0 && cell?.Value.Kind == CellValueKind.Number) style = style with { Horizontal = 3 };
                    style = style with { FontName = style.FontName ?? grid.CellFontName };
                    using var recorder = new SKPictureRecorder();
                    var target = recorder.BeginRecording(SKRect.Create((float)rect.Width + 1, (float)rect.Height + 1));
                    Renderer.DrawCell(target, new(new(0, 0, rect.Width, rect.Height),
                        new(display.Text, style, RichText: cell?.Content.RichText, DefaultFontSize: grid.CellFontSize,
                            FillOffset: display.FillOffset, FillCharacter: display.FillCharacter),
                        _sheet.ShowGridLines, _sheet.Annotations.Notes.ContainsKey(new(range.Top, range.Left))), RenderOutput.Screen,
                        location: $"{_sheet.Name}!{range.Top + 1}:{range.Left + 1}");
                    content = recorder.EndRecording();
                    // Native recordings are bounded by both count and byte estimate, including huge cells.
                    if (_cellDrawings.Count >= 4096 || _pictureBytes + content.ApproximateBytesUsed > 32L * 1024 * 1024) ClearPictures();
                    _cellDrawings.Add(key, content); _pictureBytes += content.ApproximateBytesUsed; CellDrawingCount++;
                    RecordMilliseconds += Stopwatch.GetElapsedTime(recording).TotalMilliseconds;
                }
                var small = rect.Width * sx <= 4096 && rect.Height * sy <= 4096 && rect.Width * sx * rect.Height * sy <= 1_048_576;
                if (small)
                {
                    if (!_cellRasters.TryGetValue(key, out var raster))
                    {
                        var rasterizing = Stopwatch.GetTimestamp();
                        raster = Rasterize(content, new(0, 0, rect.Width, rect.Height), sx, sy);
                        RasterMilliseconds += Stopwatch.GetElapsedTime(rasterizing).TotalMilliseconds;
                        if (_rasterBytes + Bytes(raster) > 64L * 1024 * 1024) { _cellRasters.Clear(); _chartRasters.Clear(); _rasterBytes = 0; }
                        _cellRasters[key] = raster; _rasterBytes += Bytes(raster);
                    }
                    drawing.DrawImage(raster, rect);
                }
                else DrawCropped(content, rect, clip);
            });
        }
        foreach (var row in rows)
        foreach (var column in columns)
        {
            var merge = grid.GetMergedModelRange(row, column);
            if (merge.Top != merge.Bottom || merge.Left != merge.Right) continue;
            PaintCell(new(row, column, row, column));
        }
        foreach (var merge in _merges) PaintCell(merge);
        foreach (var key in _cellDrawings.Keys.Where(key => !used.Contains(key)).ToArray())
        {
            _pictureBytes -= _cellDrawings[key].ApproximateBytesUsed; _cellDrawings[key].Dispose(); _cellDrawings.Remove(key);
            if (_cellRasters.Remove(key, out var raster)) _rasterBytes -= Bytes(raster);
        }
        void DrawCropped(SKPicture picture, Rect rect, Rect clip)
        {
            reusable = false; // A viewport-sized crop must be regenerated when it moves.
            var visible = Rect.Intersect(rect, clip); if (visible.IsEmpty) return;
            // Snap the local crop to device pixels; it is always bounded by the viewport.
            var x = Math.Floor((visible.Left - rect.Left) * sx) / sx; var y = Math.Floor((visible.Top - rect.Top) * sy) / sy;
            var crop = new Rect(x, y, Math.Ceiling((visible.Right - rect.Left - x) * sx) / sx, Math.Ceiling((visible.Bottom - rect.Top - y) * sy) / sy);
            if (crop.Width * sx * crop.Height * sy * 4 > Renderer.Budget.MaxPreviewBytes) return;
            var raster = Rasterize(picture, crop, sx, sy);
            drawing.DrawImage(raster, new(rect.Left + x, rect.Top + y, crop.Width, crop.Height));
        }
        void PaintAttachment(DrawingAnchor anchor, Action<Rect, Rect> paint)
        {
            var bounds = anchor.Bounds(_sheet); if (bounds.Width <= 0 || bounds.Height <= 0) return;
            var cells = DrawingGeometry.FromBounds(_sheet, bounds, DrawingAnchorKind.TwoCell);
            var range = new CellRectangle(cells.Row, cells.Column, Math.Min(_surface.RowCount - 1, cells.EndRow), Math.Min(_surface.ColumnCount - 1, cells.EndColumn));
            if (range.Top > range.Bottom || range.Left > range.Right) return;
            Paint(range, (rect, clip) => paint(new(rect.Left + bounds.X - DrawingGeometry.ColumnStart(_sheet, range.Left),
                rect.Top + bounds.Y - DrawingGeometry.RowStart(_sheet, range.Top), bounds.Width, bounds.Height), clip));
        }
        foreach (var image in _sheet.Images)
            PaintAttachment(image.Anchor, (rect, clip) =>
            {
                usedImages.Add(image);
                if (!_images.TryGetValue(image, out var decoded))
                {
                    try { using var pixels = Renderer.DecodeImage(image); decoded = SkiaWpfBridge.Copy(pixels); }
                    catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentException) { decoded = null; }
                    if (decoded is not null && _imageBytes + Bytes(decoded) > Renderer.Budget.MaxImageCacheBytes)
                    { _images.Clear(); _imageBytes = 0; }
                    _images[image] = decoded; _imageBytes += Bytes(decoded);
                }
                if (decoded is not null) drawing.DrawImage(decoded, rect);
                else
                {
                    using var recorder = new SKPictureRecorder();
                    Renderer.DrawText(recorder.BeginRecording(SKRect.Create((float)rect.Width, (float)rect.Height)), new("图片无法预览", new()), SKRect.Create((float)rect.Width, (float)rect.Height));
                    using var picture = recorder.EndRecording(); DrawCropped(picture, rect, clip);
                }
            });
        foreach (var chart in _sheet.Data.Charts)
            PaintAttachment(chart.Anchor, (rect, clip) =>
            {
                usedCharts.Add(chart);
                if (!_charts.TryGetValue(chart, out var picture))
                {
                    using var recorder = new SKPictureRecorder();
                    var target = recorder.BeginRecording(SKRect.Create((float)rect.Width, (float)rect.Height));
                    try { Renderer.DrawChart(target, new(0, 0, rect.Width, rect.Height), _surface.Snapshot, chart); }
                    catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException or ArgumentException or FormatException)
                    { Renderer.DrawText(target, new("图表无法预览：" + ex.Message, new()), SKRect.Create((float)rect.Width, (float)rect.Height)); }
                    picture = recorder.EndRecording();
                    if (_charts.Count >= 4096 || _pictureBytes + picture.ApproximateBytesUsed > 32L * 1024 * 1024) ClearPictures();
                    _charts[chart] = picture; _pictureBytes += picture.ApproximateBytesUsed;
                }
                if (rect.Width * sx * rect.Height * sy <= 1_048_576)
                {
                    var key = (chart, sx, sy);
                    if (!_chartRasters.TryGetValue(key, out var raster))
                    {
                        raster = Rasterize(picture, new(0, 0, rect.Width, rect.Height), sx, sy);
                        if (_rasterBytes + Bytes(raster) > 64L * 1024 * 1024) { _cellRasters.Clear(); _chartRasters.Clear(); _rasterBytes = 0; }
                        _chartRasters[key] = raster; _rasterBytes += Bytes(raster);
                    }
                    drawing.DrawImage(raster, rect);
                }
                else DrawCropped(picture, rect, clip);
            });
        foreach (var key in _images.Keys.Where(key => !usedImages.Contains(key)).ToArray())
        { _imageBytes -= Bytes(_images[key]); _images.Remove(key); }
        foreach (var key in _charts.Keys.Where(key => !usedCharts.Contains(key)).ToArray())
        { _pictureBytes -= _charts[key].ApproximateBytesUsed; _charts[key].Dispose(); _charts.Remove(key); }
        foreach (var key in _chartRasters.Keys.Where(key => !usedCharts.Contains(key.Chart) || key.ScaleX != sx || key.ScaleY != sy).ToArray())
        { _rasterBytes -= Bytes(_chartRasters[key]); _chartRasters.Remove(key); }
        _surface.ReportFormattingDiagnostics();
        recordingPanes.Dispose();
        var panes = recordingPanes.Items.Values.Select(item => item.Pane).ToArray();
        recordingPanes.Items.Clear();
        foreach (var pane in panes) pane.Drawing.Freeze();
        DrawPanes(panes);
        if (reusable)
        {
            _panes = panes; _sceneSize = size; _sceneRows = rows; _sceneColumns = columns;
            _sceneHeights = heights; _sceneWidths = widths;
        }
        RenderMilliseconds += Stopwatch.GetElapsedTime(started).TotalMilliseconds;
        RenderingTrace?.Invoke("artwork-end", Stopwatch.GetTimestamp());
    }
    private static long Bytes(BitmapSource? bitmap) => bitmap is null ? 0 : (long)bitmap.PixelWidth * bitmap.PixelHeight * 4;
    private static BitmapSource Rasterize(SKPicture picture, Rect crop, double sx, double sy)
    {
        using var pixels = new SKBitmap(Math.Max(1, (int)Math.Ceiling(crop.Width * sx)), Math.Max(1, (int)Math.Ceiling(crop.Height * sy)), SKColorType.Bgra8888, SKAlphaType.Premul);
        using var canvas = new SKCanvas(pixels); canvas.Clear(SKColors.Transparent); canvas.Scale((float)sx, (float)sy);
        canvas.Translate((float)-crop.X, (float)-crop.Y); canvas.DrawPicture(picture); canvas.Flush();
        return SkiaWpfBridge.Copy(pixels, 96 * sx, 96 * sy);
    }
    // These WPF adapters configure the native editor only, never content measurement/drawing.
    internal static double FontPixels(StyleDefinition style, double defaultFontSize = 0) => style.FontSize is { } points ? points * 96 / 72 : defaultFontSize > 0 ? defaultFontSize : 11.0 * 96 / 72;
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
