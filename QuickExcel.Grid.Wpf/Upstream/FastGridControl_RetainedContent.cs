// QuickExcel opt-in content tiles. Reuses the upstream cell rasterizer and glyphs.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FastWpfGrid;

partial class FastGridControl
{
    /// <summary>The host paints all workbook cell content in the decoration layer.</summary>
    public bool ExternalCellContent { get; set; }
    private readonly record struct ContentTileKey(FastGridCellAddress Address, int Width, int Height, int CropX, int CropY, int PixelWidth, int PixelHeight);
    private readonly Dictionary<ContentTileKey, BitmapSource> _contentTiles = new();
    private WriteableBitmap _tileCanvas;
    private bool RetainedContentEnabled => UseSmoothScrolling && UseSelectionOverlay;

    private void RenderRetainedContent()
    {
        if (_isInvalidatedAll) _contentTiles.Clear();
        else
            foreach (var key in _contentTiles.Keys.ToArray())
            {
                var a = key.Address;
                if (a.IsGridHeader && _InvalidatedGridHeader ||
                    a.Row.HasValue && _invalidatedRows.Contains(a.Row.Value) ||
                    a.Column.HasValue && _invalidatedColumns.Contains(a.Column.Value) ||
                    a.IsCell && _invalidatedCells.Contains(Tuple.Create(a.Row.Value, a.Column.Value)) ||
                    a.IsRowHeader && _invalidatedRowHeaders.Contains(a.Row.Value) ||
                    a.IsColumnHeader && _invalidatedColumnHeaders.Contains(a.Column.Value)) _contentTiles.Remove(key);
            }

        var used = new HashSet<ContentTileKey>();
        var group = new DrawingGroup();
        var width = _drawBuffer.PixelWidth; var height = _drawBuffer.PixelHeight;
        var dx = _renderScaleX; var dy = _renderScaleY;
        using (var drawing = group.Open())
        {
            drawing.DrawRectangle(Brushes.White, null, new Rect(0, 0, width / dx, height / dy));
            void Tile(FastGridCellAddress address, IntRect cell, Rect clip)
            {
                if (ExternalCellContent && address.IsCell) return;
                if (clip.IsEmpty || clip.Width <= 0 || clip.Height <= 0) return;
                var rect = new Rect(cell.Left, cell.Top, cell.Width, cell.Height);
                if (!rect.IntersectsWith(clip)) return;
                var visible = rect;
                // Workbook dimensions may be much larger than the viewport. Do
                // not allocate a full bitmap for a huge, mostly off-screen cell.
                if ((long)cell.Width * cell.Height > 1_048_576 || cell.Width > 4096 || cell.Height > 4096) visible.Intersect(clip);
                var key = new ContentTileKey(address, cell.Width, cell.Height, (int)visible.Left - cell.Left,
                    (int)visible.Top - cell.Top, (int)visible.Width, (int)visible.Height);
                used.Add(key);
                if (!_contentTiles.TryGetValue(key, out var bitmap))
                {
                    bitmap = RasterizeContentTile(key, cell);
                    _contentTiles[key] = bitmap;
                }
                drawing.DrawImage(bitmap, new Rect(visible.X / dx, visible.Y / dy, visible.Width / dx, visible.Height / dy));
            }
            void PushClip(Rect clip) => drawing.PushClip(new RectangleGeometry(new Rect(clip.X / dx, clip.Y / dy, clip.Width / dx, clip.Height / dy)));
            // Draw each frozen/scrolling quadrant under its own clip. A partly
            // scrolled cell must never overwrite a frozen cell or its header.
            foreach (var rows in VisibleAxis(_rowSizes, FirstVisibleRowScrollIndex, VisibleRowCount, HeaderHeight, height))
            foreach (var columns in VisibleAxis(_columnSizes, FirstVisibleColumnScrollIndex, VisibleColumnCount, HeaderWidth, width))
            {
                var clip = new Rect(columns.Min, rows.Min, columns.Max - columns.Min, rows.Max - rows.Min);
                PushClip(clip);
                for (var r = rows.First; r <= rows.Last && r < _realRowCount; r++)
                for (var c = columns.First; c <= columns.Last && c < _realColumnCount; c++)
                    Tile(new(r, c), GetCellRect(r, c), clip);
                drawing.Pop();
            }
            foreach (var rows in VisibleAxis(_rowSizes, FirstVisibleRowScrollIndex, VisibleRowCount, HeaderHeight, height))
            {
                var clip = new Rect(0, rows.Min, Math.Min(HeaderWidth + 1, width), rows.Max - rows.Min); PushClip(clip);
                for (var r = rows.First; r <= rows.Last && r < _realRowCount; r++)
                    Tile(new(r, null), GetRowHeaderRect(r), clip);
                drawing.Pop();
            }
            foreach (var columns in VisibleAxis(_columnSizes, FirstVisibleColumnScrollIndex, VisibleColumnCount, HeaderWidth, width))
            {
                var clip = new Rect(columns.Min, 0, columns.Max - columns.Min, Math.Min(HeaderHeight + 1, height)); PushClip(clip);
                for (var c = columns.First; c <= columns.Last && c < _realColumnCount; c++)
                    Tile(new(null, c), GetColumnHeaderRect(c), clip);
                drawing.Pop();
            }
            if (Model != null) Tile(FastGridCellAddress.GridHeader, GetGridHeaderRect(), new Rect(0, 0, width, height));
        }
        // Retain only this viewport, so a long drag cannot accumulate the sheet.
        foreach (var key in _contentTiles.Keys.ToArray()) if (!used.Contains(key)) _contentTiles.Remove(key);
        group.Freeze();
        var frame = new DrawingImage(group); frame.Freeze(); image.Source = frame;
        _frameChanged = false;
        ClearInvalidation();
    }

    private BitmapSource RasterizeContentTile(ContentTileKey key, IntRect rect)
    {
        var address = key.Address;
        // Keep the original cell drawing semantics, but use a private reusable
        // scratch canvas. Only immutable small cell images reach the compositor.
        if (_tileCanvas == null || _tileCanvas.PixelWidth < key.PixelWidth || _tileCanvas.PixelHeight < key.PixelHeight)
        {
            var old = _tileCanvas;
            var w = Math.Max(key.PixelWidth, old?.PixelWidth ?? 0); var h = Math.Max(key.PixelHeight, old?.PixelHeight ?? 0);
            if ((long)w * h > Math.Max(1_048_576L, (long)key.PixelWidth * key.PixelHeight)) { w = key.PixelWidth; h = key.PixelHeight; }
            _tileCanvas = BitmapFactory.New(w, h, _renderScaleX, _renderScaleY);
            old?.Freeze();
        }
        var viewportCanvas = _drawBuffer; var pointer = _mouseCursorPoint;
        var isHover = address == _mouseOverCell;
        _drawBuffer = _tileCanvas;
        if (pointer.HasValue) _mouseCursorPoint = new Point(pointer.Value.X - rect.Left - key.CropX, pointer.Value.Y - rect.Top - key.CropY);
        try
        {
            IFastGridCell cell; Color background;
            if (address.IsCell) { cell = GetCell(address.Row.Value, address.Column.Value); background = cell?.BackgroundColor ?? GetAlternateBackground(address.Row.Value); }
            else if (address.IsRowHeader) { cell = GetRowHeader(address.Row.Value); background = cell?.BackgroundColor ?? HeaderBackground; }
            else if (address.IsColumnHeader) { cell = GetColumnHeader(address.Column.Value); background = cell?.BackgroundColor ?? HeaderBackground; }
            else { cell = Model.GetGridHeader(this); background = HeaderBackground; }
            using (var canvas = _tileCanvas.GetBitmapContext())
            {
                _tileCanvas.Clear(Colors.White);
                RenderCell(cell, new IntRect(new IntPoint(-key.CropX, -key.CropY), rect.Size), null, background, address);
                var bytes = checked((key.PixelHeight - 1) * _tileCanvas.BackBufferStride + key.PixelWidth * 4);
                var bitmap = BitmapSource.Create(key.PixelWidth, key.PixelHeight, _tileCanvas.DpiX, _tileCanvas.DpiY,
                    _tileCanvas.Format, null, _tileCanvas.BackBuffer, bytes, _tileCanvas.BackBufferStride);
                bitmap.Freeze();
                return bitmap;
            }
        }
        finally
        {
            if (isHover)
                foreach (var region in CurrentCellActiveRegions)
                    region.Rect = new IntRect(new IntPoint(region.Rect.Left + rect.Left + key.CropX, region.Rect.Top + rect.Top + key.CropY), region.Rect.Size);
            _drawBuffer = viewportCanvas; _mouseCursorPoint = pointer;
        }
    }
}
