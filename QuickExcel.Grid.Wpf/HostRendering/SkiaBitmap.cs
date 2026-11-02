using System;
using System.Collections.Generic;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace System.Windows.Media.Imaging;

// WPF host adapter only. No GDI glyphs, unmanaged CRT imports or shared mutable compositor surfaces.
public static class DpiDetector
{
    private static DpiScale Dpi => Application.Current?.MainWindow is { } window ? VisualTreeHelper.GetDpi(window) : new(1, 1);
    public static double DpiXKoef => Dpi.DpiScaleX;
    public static double DpiYKoef => Dpi.DpiScaleY;
}
public static class BitmapFactory
{
    public static WriteableBitmap New(int width, int height, double scaleX = 0, double scaleY = 0) =>
        new(Math.Max(1, width), Math.Max(1, height), (scaleX > 0 ? scaleX : DpiDetector.DpiXKoef) * 96,
            (scaleY > 0 ? scaleY : DpiDetector.DpiYKoef) * 96, PixelFormats.Pbgra32, null);
}
public enum ReadWriteMode { ReadOnly, ReadWrite }
public sealed unsafe class BitmapContext : IDisposable
{
    private readonly WriteableBitmap _bitmap;
    private readonly ReadWriteMode _mode;
    private bool _disposed;
    public int* Pixels => (int*)_bitmap.BackBuffer;
    public int Width => _bitmap.PixelWidth;
    public int Height => _bitmap.PixelHeight;
    public BitmapContext(WriteableBitmap bitmap, ReadWriteMode mode) { _bitmap = bitmap; _mode = mode; bitmap.Lock(); }
    public void Dispose()
    {
        if (_disposed) return; _disposed = true;
        if (_mode == ReadWriteMode.ReadWrite) _bitmap.AddDirtyRect(new(0, 0, Width, Height));
        _bitmap.Unlock();
    }
}
public sealed class GridFont
{
    private readonly SkiaRenderer _renderer;
    internal readonly PortableFontDesc Description;
    private readonly Dictionary<string, TextMetrics> _metrics = new();
    internal GridFont(SkiaRenderer renderer, PortableFontDesc description) { _renderer = renderer; Description = description; }
    internal TextLayoutRequest Request(string text, Color? color = null) => new(text ?? "", new()
    {
        FontName = Description.FontName, FontSize = Description.EmSize * .75, Bold = Description.IsBold, Italic = Description.IsItalic,
        FontColor = color is { } c ? $"{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}" : null, WrapText = true
    });
    internal SkiaTextLayout Layout(string text, Color color) => _renderer.Layout(Request(text, color));
    private TextMetrics Measure(string text)
    {
        text ??= "";
        if (_metrics.TryGetValue(text, out var metrics)) return metrics;
        using var layout = _renderer.Layout(Request(text)); metrics = layout.Metrics;
        if (_metrics.Count >= 4096) _metrics.Clear();
        _metrics[text] = metrics; return metrics;
    }
    public int GetTextWidth(string text, int? maxSize = null) => Math.Min(maxSize ?? int.MaxValue, (int)Math.Ceiling(Measure(text).Width * Description.ScaleX));
    public int GetTextHeight(string text) => (int)Math.Ceiling(Measure(text).Height * Description.ScaleY);
    public int TextHeight => GetTextHeight("Ag");
}
public static partial class WriteableBitmapExtensions
{
    public enum BlendMode { Alpha, ColorKeying }
    public static BitmapContext GetBitmapContext(this WriteableBitmap bitmap, ReadWriteMode mode = ReadWriteMode.ReadWrite) => new(bitmap, mode);
    private static SKColor ColorOf(Color c) => new(c.R, c.G, c.B, c.A);
    private static SKImageInfo Info(WriteableBitmap bitmap) => new(bitmap.PixelWidth, bitmap.PixelHeight, SKColorType.Bgra8888,
        bitmap.Format == PixelFormats.Bgra32 ? SKAlphaType.Unpremul : SKAlphaType.Premul);
    internal static void Paint(WriteableBitmap bitmap, Action<SKCanvas> paint)
    {
        using var context = bitmap.GetBitmapContext();
        using var surface = SKSurface.Create(Info(bitmap), bitmap.BackBuffer, bitmap.BackBufferStride);
        paint(surface.Canvas); surface.Canvas.Flush();
    }
    public static void Clear(this WriteableBitmap bitmap, Color color) => Paint(bitmap, canvas => canvas.Clear(ColorOf(color)));
    public static void FillRectangle(this WriteableBitmap bitmap, int left, int top, int right, int bottom, Color color) => Paint(bitmap, canvas =>
    { using var paint = new SKPaint { Color = ColorOf(color) }; canvas.DrawRect(new(left, top, right, bottom), paint); });
    public static void DrawRectangle(this WriteableBitmap bitmap, int left, int top, int right, int bottom, Color color) => Paint(bitmap, canvas =>
    { using var paint = new SKPaint { Color = ColorOf(color), Style = SKPaintStyle.Stroke, StrokeWidth = 1 }; canvas.DrawRect(new(left + .5f, top + .5f, right + .5f, bottom + .5f), paint); });
    public static void DrawLine(this WriteableBitmap bitmap, int x1, int y1, int x2, int y2, Color color) => Paint(bitmap, canvas =>
    { using var paint = new SKPaint { Color = ColorOf(color), StrokeWidth = 1 }; canvas.DrawLine(x1 + .5f, y1 + .5f, x2 + .5f, y2 + .5f, paint); });
    public static int DrawString(this WriteableBitmap bitmap, int x, int y, IntRect clip, Color color, Color? background, GridFont font, string text)
    {
        if (string.IsNullOrEmpty(text)) return 0;
        using var layout = font.Layout(text, color);
        Paint(bitmap, canvas =>
        {
            canvas.ClipRect(new(clip.Left, clip.Top, clip.Right + 1, clip.Bottom + 1)); canvas.Translate(x, y);
            canvas.Scale((float)font.Description.ScaleX, (float)font.Description.ScaleY); layout.Draw(canvas, 0, 0);
        });
        return font.GetTextWidth(text);
    }
    public static unsafe Color GetPixel(this WriteableBitmap bitmap, int x, int y)
    {
        using var context = bitmap.GetBitmapContext(ReadWriteMode.ReadOnly);
        var value = *(uint*)((byte*)bitmap.BackBuffer + y * bitmap.BackBufferStride + x * 4);
        return Color.FromArgb((byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value);
    }
    public static void Blit(this WriteableBitmap bitmap, Point point, WriteableBitmap source, Rect sourceRect, Color key, BlendMode mode)
    {
        using var context = source.GetBitmapContext(ReadWriteMode.ReadOnly);
        using var pixels = new SKBitmap(); pixels.InstallPixels(Info(source), source.BackBuffer, source.BackBufferStride);
        using var copy = pixels.Copy();
        if (mode == BlendMode.ColorKeying)
            for (var y = 0; y < copy.Height; y++) for (var x = 0; x < copy.Width; x++)
                if (copy.GetPixel(x, y) == ColorOf(key)) copy.SetPixel(x, y, SKColors.Transparent);
        Paint(bitmap, canvas =>
        {
            using var image = SKImage.FromBitmap(copy);
            canvas.DrawImage(image, SKRect.Create((float)sourceRect.X, (float)sourceRect.Y, (float)sourceRect.Width, (float)sourceRect.Height),
                SKRect.Create((float)point.X, (float)point.Y, (float)sourceRect.Width, (float)sourceRect.Height), new SKSamplingOptions(SKFilterMode.Nearest));
        });
    }
    public static void ScrollX(this WriteableBitmap bitmap, int dx, IntRect rect) => Scroll(bitmap, dx, 0, rect);
    public static void ScrollY(this WriteableBitmap bitmap, int dy, IntRect rect) => Scroll(bitmap, 0, dy, rect);
    private static unsafe void Scroll(WriteableBitmap bitmap, int dx, int dy, IntRect rect)
    {
        var left = Math.Max(0, rect.Left); var right = Math.Min(bitmap.PixelWidth, rect.Right + 1);
        var top = Math.Max(0, rect.Top); var bottom = Math.Min(bitmap.PixelHeight, rect.Bottom + 1);
        if (right <= left || bottom <= top) return;
        using var context = bitmap.GetBitmapContext();
        var stride = bitmap.BackBufferStride / 4;
        for (var i = 0; i < bottom - top; i++)
        {
            var y = dy > 0 ? bottom - 1 - i : top + i; var sourceY = y - dy;
            var row = new Span<int>(context.Pixels + y * stride + left, right - left);
            if (sourceY < top || sourceY >= bottom || Math.Abs(dx) >= row.Length) { row.Fill(unchecked((int)0xffffffff)); continue; }
            var source = new Span<int>(context.Pixels + sourceY * stride + left, row.Length);
            if (dx > 0) { source[..^dx].CopyTo(row[dx..]); row[..dx].Fill(unchecked((int)0xffffffff)); }
            else if (dx < 0) { source[-dx..].CopyTo(row[..(row.Length + dx)]); row[(row.Length + dx)..].Fill(unchecked((int)0xffffffff)); }
            else source.CopyTo(row);
        }
    }
}
