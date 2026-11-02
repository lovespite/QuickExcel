// QuickExcel host boundary: device scale of the raster viewport.
// Content pixels are device pixels of the *final* composition. Hosts apply
// worksheet zoom (or any ancestor scale) as a layout transform, so the legacy
// window DPI alone under-samples the bitmap and every raster glyph/border is
// resampled at a fractional scale. The control measures its own transform to
// the compositor and rasterizes at that scale instead.
using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FastWpfGrid;

partial class FastGridControl
{
    private double _renderScaleX = 1;
    private double _renderScaleY = 1;
    private bool _syncingRenderScale;

    /// <summary>Device pixels per DIP horizontally for the current composition.</summary>
    public double RenderScaleX => _renderScaleX;
    /// <summary>Device pixels per DIP vertically for the current composition.</summary>
    public double RenderScaleY => _renderScaleY;

    private static (double X, double Y) MeasureDeviceScale(Visual visual)
    {
        var dpi = VisualTreeHelper.GetDpi(visual);
        var x = dpi.DpiScaleX;
        var y = dpi.DpiScaleY;
        try
        {
            var source = PresentationSource.FromVisual(visual);
            if (source?.CompositionTarget != null && source.RootVisual is Visual root &&
                !ReferenceEquals(root, visual) && root.IsAncestorOf(visual))
            {
                // Ancestor layout transforms (worksheet zoom) are part of this
                // transform; the composition target adds the monitor scale.
                var toRoot = visual.TransformToAncestor(root);
                var origin = toRoot.Transform(new Point(0, 0));
                var unit = toRoot.Transform(new Point(1, 1));
                var scaleX = (unit.X - origin.X) * dpi.DpiScaleX;
                var scaleY = (unit.Y - origin.Y) * dpi.DpiScaleY;
                if (scaleX > 0.01 && scaleX < 100) x = scaleX;
                if (scaleY > 0.01 && scaleY < 100) y = scaleY;
            }
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException)
        {
            // Not connected to a presentation source yet; the DPI fallback is
            // replaced as soon as the control is loaded or resized.
        }
        return (x, y);
    }

    /// <summary>Re-measures the device scale; returns true when the viewport must be rebuilt.</summary>
    private bool UpdateRenderScale()
    {
        if (_syncingRenderScale) return false;
        var (x, y) = MeasureDeviceScale(this);
        if (Math.Abs(x - _renderScaleX) < 1e-9 && Math.Abs(y - _renderScaleY) < 1e-9) return false;
        _syncingRenderScale = true;
        try
        {
            _renderScaleX = x;
            _renderScaleY = y;
            // Glyph bitmaps and retained cell tiles embed the scale, and stored
            // dimensions are in content pixels, so all three are rebuilt.
            ClearCaches();
            _contentTiles.Clear();
            _tileCanvas?.Freeze();
            _tileCanvas = null;
            ApplyStoredDimensions();
            RecalculateHeaderSize();
            RecalculateDefaultCellSizeIfUnowned();
            SetScrollbarMargin();
        }
        finally
        {
            _syncingRenderScale = false;
        }
        return true;
    }

    private void RecalculateDefaultCellSizeIfUnowned()
    {
        // Explicit workbook dimensions win; only the text auto-fit path derives
        // default sizes from the (now scaled) glyph metrics.
        if (!_hasModelDimensions) RecalculateDefaultCellSize();
    }

    /// <summary>Resizes the software viewport while keeping its device density tag.</summary>
    private static WriteableBitmap CropBuffer(WriteableBitmap source, int width, int height)
    {
        var result = BitmapFactory.New(width, height, source.DpiX / 96.0, source.DpiY / 96.0);
        var copyWidth = Math.Min(width, source.PixelWidth);
        var copyHeight = Math.Min(height, source.PixelHeight);
        if (copyWidth <= 0 || copyHeight <= 0) return result;
        var sourceStride = source.BackBufferStride / 4;
        var destinationStride = result.BackBufferStride / 4;
        using var src = source.GetBitmapContext(ReadWriteMode.ReadOnly);
        using var dst = result.GetBitmapContext();
        unsafe
        {
            for (var y = 0; y < copyHeight; y++)
                Buffer.MemoryCopy(src.Pixels + (long)y * sourceStride, dst.Pixels + (long)y * destinationStride,
                    (long)copyWidth * 4, (long)copyWidth * 4);
        }
        return result;
    }

    protected override void OnDpiChanged(DpiScale oldDpi, DpiScale newDpi)
    {
        base.OnDpiChanged(oldDpi, newDpi);
        if (!IsLoaded) return;
        if (UpdateRenderScale()) RebuildViewport();
    }
}
