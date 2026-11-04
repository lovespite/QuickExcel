using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private readonly record struct PictureKey(GridCellRange Range, double Width, double Height, bool Header = false);
    private readonly record struct Fragment(Rect Bounds, Rect Clip);
    private readonly Dictionary<PictureKey, SKPicture> _pictures = [];
    private readonly Dictionary<SheetImage, SKBitmap?> _images = [];
    private readonly Dictionary<WorksheetChart, SKPicture> _charts = [];
    private long _pictureBytes, _imageBytes;
    public long CellDrawingCount { get; private set; }
    public long ContentRenderCount { get; private set; }
    public double LastContentRenderMilliseconds { get; private set; }
    public long CachedNativeBytes => _pictureBytes + _imageBytes;
    public event Action<string, long>? RenderingTrace;
    private string? FilterHeader(int row, int column)
    {
        if (Sheet.Data.Filter is { } filter && filter.Range.Top == row && column >= filter.Range.Left && column <= filter.Range.Right && filter.Columns.GetValueOrDefault(column)?.ShowButton != false)
            return SheetViewState.HeaderMarker(column, filter, Sheet.Data.Sort);
        foreach (var table in Sheet.Data.Tables)
            if (table.Filter is { } candidate && candidate.Range.Top == row && column >= candidate.Range.Left && column <= candidate.Range.Right && candidate.Columns.GetValueOrDefault(column)?.ShowButton != false)
                return SheetViewState.HeaderMarker(column, candidate, table.Sort);
        return null;
    }

    private void ClearPictures()
    {
        foreach (var picture in _pictures.Values) picture.Dispose();
        foreach (var chart in _charts.Values) chart.Dispose();
        _pictures.Clear(); _charts.Clear(); _pictureBytes = 0;
    }
    private void ClearScene()
    {
        ClearPictures(); _artwork.Release();
        foreach (var image in _images.Values) image?.Dispose();
        _images.Clear(); _imageBytes = 0;
    }
    private void ReleaseRendering() { ClearScene(); _renderer?.Dispose(); _renderer = null; }
    private IEnumerable<Fragment> Fragments(GridCellRange range)
    {
        foreach (var row in _rows.Project(range.Top, range.Bottom, ViewHeight))
        foreach (var column in _columns.Project(range.Left, range.Right, ViewWidth))
            yield return new(new(HeaderWidth + column.Start, HeaderHeight + row.Start, column.Size, row.Size),
                new(HeaderWidth + column.ClipStart, HeaderHeight + row.ClipStart, column.ClipSize, row.ClipSize));
    }
    private GridCellRange CellRange(CellPosition position) => _merges.FirstOrDefault(m => m.Contains(position.Row, position.Column)) is var merge && merge != default
        ? merge : new(position.Row, position.Column, position.Row, position.Column);
    private static SKRect SkiaRect(Rect rect) => SKRect.Create((float)rect.X, (float)rect.Y, (float)rect.Width, (float)rect.Height);
    private void Paint(SKCanvas canvas)
    {
        if (_snapshot is null) return;
        var start = Stopwatch.GetTimestamp(); RenderingTrace?.Invoke("content-start", start);
        canvas.Clear(SKColors.White); canvas.Scale((float)_zoom);
        var rows = _rows.Visible(ViewHeight); var columns = _columns.Visible(ViewWidth);
        var drawn = new HashSet<GridCellRange>();
        void Picture(GridCellRange range)
        {
            if (!drawn.Add(range)) return;
            foreach (var fragment in Fragments(range))
            {
                var bounds = fragment.Bounds;
                var key = new PictureKey(range, bounds.Width, bounds.Height);
                if (!_pictures.TryGetValue(key, out var picture))
                {
                    var cell = Sheet.GetCell(range.Top, range.Left);
                    var style = ResolveStyle?.Invoke(_snapshot, Sheet, range.Top, range.Left) ?? _snapshot.Styles[Sheet.GetStyleId(range.Top, range.Left)];
                    var display = CellDisplayFormatter.FormatDisplay(cell?.Value ?? CellValue.Blank, style, _snapshot.DateSystem);
                    if (display.Color is not null) style = style with { FontColor = display.Color };
                    if (style.Horizontal == 0 && cell?.Value.Kind == CellValueKind.Number) style = style with { Horizontal = 3 };
                    using var recorder = new SKPictureRecorder(); var target = recorder.BeginRecording(SKRect.Create((float)bounds.Width + 1, (float)bounds.Height + 1));
                    Renderer.DrawCell(target, new(new(0, 0, bounds.Width, bounds.Height),
                        new(display.Text, style, RichText: cell?.Content.RichText, DefaultFontSize: 14, FillOffset: display.FillOffset, FillCharacter: display.FillCharacter),
                        Sheet.ShowGridLines, Sheet.Annotations.Notes.ContainsKey(new(range.Top, range.Left))), RenderOutput.Screen);
                    if (FilterHeader(range.Top, range.Left) is { } marker)
                        Renderer.DrawCell(target, new(new(Math.Max(0, bounds.Width - 22), 0, 22, bounds.Height),
                            new(marker, new() { FillColor = "FFF2F5F4", FontColor = "FF006B59", Horizontal = 2, Vertical = 1 }, DefaultFontSize: 10)), RenderOutput.Screen);
                    picture = recorder.EndRecording(); AdmitPicture(picture); _pictures[key] = picture; CellDrawingCount++;
                }
                canvas.Save(); canvas.ClipRect(SkiaRect(fragment.Clip)); canvas.Translate((float)bounds.X, (float)bounds.Y); canvas.DrawPicture(picture); canvas.Restore();
            }
        }
        foreach (var row in rows)
        foreach (var col in columns) Picture(CellRange(new(row.ModelIndex, col.ModelIndex)));
        PaintChartSources(canvas);
        PaintAttachments(canvas);
        var header = new StyleDefinition { FillColor = "FFF2F5F4", FontColor = "FF42534E", Horizontal = 2, Vertical = 1 };
        var rowBreaks = Sheet.Printing.RowBreaks.Where(b => b.Manual).Select(b => b.Id).ToHashSet();
        var columnBreaks = Sheet.Printing.ColumnBreaks.Where(b => b.Manual).Select(b => b.Id).ToHashSet();
        void Header(int row, int column, Rect bounds, Rect clip, string text)
        {
            var key = new PictureKey(new(row, column, row, column), bounds.Width, bounds.Height, true);
            if (!_pictures.TryGetValue(key, out var picture))
            {
                using var recorder = new SKPictureRecorder();
                var target = recorder.BeginRecording(SKRect.Create((float)bounds.Width + 1, (float)bounds.Height + 1));
                Renderer.DrawCell(target,
                    new(new(0, 0, bounds.Width, bounds.Height), new(text, header, DefaultFontSize: 12), true), RenderOutput.Screen);
                if (column == -1 && rowBreaks.Contains(row) || row == -1 && columnBreaks.Contains(column))
                {
                    target.Save(); target.ClipRect(SKRect.Create((float)bounds.Width, (float)bounds.Height));
                    using var paint = new SKPaint { Color = new(37, 99, 235), IsAntialias = true, StrokeWidth = 2 };
                    using var marker = new SKPathBuilder();
                    if (column == -1)
                    {
                        target.DrawLine(3, 3, 17, 3, paint);
                        marker.MoveTo(6, 3); marker.LineTo(14, 3); marker.LineTo(10, 9);
                    }
                    else
                    {
                        target.DrawLine(3, 3, 3, 17, paint);
                        marker.MoveTo(3, 6); marker.LineTo(3, 14); marker.LineTo(9, 10);
                    }
                    marker.Close(); using var path = marker.Detach(); target.DrawPath(path, paint); target.Restore();
                }
                picture = recorder.EndRecording(); AdmitPicture(picture); _pictures[key] = picture;
            }
            canvas.Save(); canvas.ClipRect(SkiaRect(clip)); canvas.Translate((float)bounds.X, (float)bounds.Y); canvas.DrawPicture(picture); canvas.Restore();
        }
        foreach (var col in columns) Header(-1, col.ModelIndex, new(HeaderWidth + col.Start, 0, col.Size, HeaderHeight),
            new(HeaderWidth + (col.Frozen ? 0 : _columns.FrozenSize), 0, col.Frozen ? _columns.FrozenSize : ViewWidth - _columns.FrozenSize, HeaderHeight), ColumnName(col.ModelIndex));
        foreach (var row in rows) Header(row.ModelIndex, -1, new(0, HeaderHeight + row.Start, HeaderWidth, row.Size),
            new(0, HeaderHeight + (row.Frozen ? 0 : _rows.FrozenSize), HeaderWidth, row.Frozen ? _rows.FrozenSize : ViewHeight - _rows.FrozenSize), (row.ModelIndex + 1).ToString());
        Header(-1, -1, new(0, 0, HeaderWidth, HeaderHeight), new(0, 0, HeaderWidth, HeaderHeight), "");
        ContentRenderCount++; LastContentRenderMilliseconds = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
        RenderingTrace?.Invoke("content-end", Stopwatch.GetTimestamp());
    }
    private void AdmitPicture(SKPicture picture)
    {
        if (picture.ApproximateBytesUsed > 32L * 1024 * 1024) { picture.Dispose(); throw new InvalidOperationException("单格绘制超过缓存预算。"); }
        if (_pictures.Count + _charts.Count >= 4096 || _pictureBytes + picture.ApproximateBytesUsed > 32L * 1024 * 1024) ClearPictures();
        _pictureBytes += picture.ApproximateBytesUsed;
    }
    private void PaintAttachments(SKCanvas canvas)
    {
        foreach (var (image, index) in Sheet.Images.Select((image, index) => (image, index)))
        foreach (var fragment in ObjectFragments(DisplayAnchor(new(DrawingObjectKind.Image, index), image.Anchor)))
        {
            if (!_images.TryGetValue(image, out var pixels))
            {
                try { pixels = Renderer.DecodeImage(image); }
                catch (Exception ex) when (ex is IOException or NotSupportedException or ArgumentException) { pixels = null; }
                var bytes = pixels is null ? 0 : (long)pixels.RowBytes * pixels.Height;
                if (_imageBytes + bytes > Renderer.Budget.MaxImageCacheBytes)
                { foreach (var old in _images.Values) old?.Dispose(); _images.Clear(); _imageBytes = 0; }
                _images[image] = pixels; _imageBytes += bytes;
            }
            canvas.Save(); canvas.ClipRect(SkiaRect(fragment.Clip));
            if (pixels is not null) SkiaRenderer.DrawImage(canvas, pixels, SkiaRect(fragment.Bounds), false, DisplayAnchor(new(DrawingObjectKind.Image, index), image.Anchor).Rotation);
            else Renderer.DrawText(canvas, new("图片无法预览", new()), SkiaRect(fragment.Bounds));
            canvas.Restore();
        }
        foreach (var (originalChart, index) in Sheet.Data.Charts.Select((chart, index) => (chart, index)))
        {
            var chart = originalChart with { Anchor = DisplayAnchor(new(DrawingObjectKind.Chart, index), originalChart.Anchor) };
            foreach (var fragment in ObjectFragments(chart.Anchor))
            {
                if (!_charts.TryGetValue(chart, out var picture))
                {
                    using var recorder = new SKPictureRecorder(); var target = recorder.BeginRecording(SKRect.Create((float)fragment.Bounds.Width, (float)fragment.Bounds.Height));
                    try { Renderer.DrawChart(target, new(0, 0, fragment.Bounds.Width, fragment.Bounds.Height), _snapshot!, chart with { Anchor = chart.Anchor with { Rotation = 0 } }); }
                    catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException or ArgumentException or FormatException)
                    { Renderer.DrawText(target, new("图表无法预览：" + ex.Message, new()), SKRect.Create((float)fragment.Bounds.Width, (float)fragment.Bounds.Height)); }
                    picture = recorder.EndRecording(); AdmitPicture(picture); _charts[chart] = picture;
                }
                canvas.Save(); canvas.ClipRect(SkiaRect(fragment.Clip)); canvas.RotateDegrees((float)chart.Anchor.Rotation, (float)fragment.Bounds.Center.X, (float)fragment.Bounds.Center.Y); canvas.Translate((float)fragment.Bounds.X, (float)fragment.Bounds.Y); canvas.DrawPicture(picture); canvas.Restore();
            }
        }
    }
    private void PaintOverlay(DrawingContext context)
    {
        if (_snapshot is null) return;
        using var transform = context.PushTransform(Matrix.CreateScale(_zoom, _zoom));
        if (_selectedObject is not null) { PaintDrawingSelection(context); return; }
        var selectedCell = Sheet.GetCell(_current.Row,_current.Column);
        var spillAnchor = selectedCell?.SpillOwner is { } owner ? new CellPosition(owner.Row,owner.Column) : _current;
        if (Sheet.GetCell(spillAnchor.Row,spillAnchor.Column)?.Calculation?.Array is { } array)
        foreach (var fragment in Fragments(new GridCellRange(spillAnchor.Row,spillAnchor.Column,
            spillAnchor.Row + array.Rows - 1,spillAnchor.Column + array.Columns - 1)))
        {
            using var clip = context.PushClip(fragment.Clip);
            context.DrawRectangle(null,new Pen(Brushes.DodgerBlue,1 / _zoom),fragment.Bounds);
        }
        foreach (var range in _selection.Ranges)
        foreach (var fragment in Fragments(range))
        {
            using var clip = context.PushClip(fragment.Clip);
            context.DrawRectangle(new SolidColorBrush(Color.FromArgb(22, 0, 128, 110)), new Pen(Brushes.Teal, 1.5 / _zoom), fragment.Bounds);
        }
        foreach (var range in new[] { _fillPreview, _referenceRange }.OfType<GridCellRange>())
        foreach (var fragment in Fragments(range))
        {
            using var clip = context.PushClip(fragment.Clip);
            context.DrawRectangle(null, new Pen(Brushes.DodgerBlue, 2 / _zoom), fragment.Bounds);
        }
        if (FillHandle() is { } handle)
            context.DrawRectangle(Brushes.Teal, null, new Rect(handle.X / _zoom, handle.Y / _zoom, handle.Width / _zoom, handle.Height / _zoom));
        RenderingTrace?.Invoke("selection-end", Stopwatch.GetTimestamp());
    }
}
