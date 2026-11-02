using System.Windows.Media;
using System.Windows.Media.Imaging;
using SkiaSharp;

namespace QuickExcel.Desktop;

/// <summary>Host-only pixel copy. WPF never holds a pointer into a mutable/reused Skia surface.</summary>
internal static class SkiaWpfBridge
{
    internal static BitmapSource Copy(SKBitmap bitmap, double dpiX = 96, double dpiY = 96)
    {
        var source = BitmapSource.Create(bitmap.Width, bitmap.Height, dpiX, dpiY, PixelFormats.Pbgra32, null,
            bitmap.GetPixels(), checked(bitmap.RowBytes * bitmap.Height), bitmap.RowBytes);
        source.Freeze(); return source;
    }
}
