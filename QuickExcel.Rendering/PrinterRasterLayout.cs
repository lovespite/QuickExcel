namespace QuickExcel.Rendering;

/// <summary>Maps the existing 96-DPI page to the physical paper, compensating for the printer's hard origin.</summary>
public sealed record PrinterRasterLayout(int X, int Y, int Width, int Height, int RasterDpi)
{
    public static PrinterRasterLayout Create(double paperWidth, double paperHeight, int dpiX, int dpiY,
        int physicalWidth, int physicalHeight, int offsetX, int offsetY)
    {
        if (dpiX is < 1 or > 100_000 || dpiY is < 1 or > 100_000 || physicalWidth <= 0 || physicalHeight <= 0 || offsetX < 0 || offsetY < 0 ||
            !double.IsFinite(paperWidth) || !double.IsFinite(paperHeight) || paperWidth <= 0 || paperHeight <= 0)
            throw new ArgumentException("打印机返回的页面几何无效。");
        var width = checked((int)Math.Round(paperWidth / 96 * dpiX)); var height = checked((int)Math.Round(paperHeight / 96 * dpiY));
        if (Math.Abs(width - physicalWidth) > dpiX / 25.4 || Math.Abs(height - physicalHeight) > dpiY / 25.4)
            throw new InvalidOperationException("打印机的纸张或方向与预览不同。请使系统打印设置与预览一致，或更改页面设置后重新生成。");
        var dpi = (int)Math.Floor(Math.Min(Math.Min(300, Math.Min(dpiX, dpiY)), Math.Sqrt((64L * 1024 * 1024 - 256 * 1024) / 4d / paperWidth / paperHeight) * 96));
        if (dpi < 24) throw new NotSupportedException("纸张尺寸超过有界栅格打印预算。");
        return new(-offsetX, -offsetY, width, height, dpi);
    }
}
