namespace QuickExcel.Rendering.Test;

public sealed class PrinterRasterLayoutTests
{
    [Fact]
    public void PaperOriginCompensatesHardMarginsWithoutRescalingTheFrozenPage()
    {
        var layout = PrinterRasterLayout.Create(816, 1056, 600, 600, 5100, 6600, 75, 100);
        Assert.Equal(-75, layout.X); Assert.Equal(-100, layout.Y); Assert.Equal(5100, layout.Width); Assert.Equal(6600, layout.Height); Assert.Equal(300, layout.RasterDpi);
        Assert.Throws<InvalidOperationException>(() => PrinterRasterLayout.Create(816, 1056, 600, 600, 6600, 5100, 75, 100));
    }
    [Fact]
    public void LargePaperReducesRasterDpiWithinTheSameMemoryBudget()
    {
        var layout = PrinterRasterLayout.Create(1584, 2448, 600, 600, 9900, 15300, 50, 50);
        Assert.InRange(layout.RasterDpi, 24, 299);
        Assert.True(Math.Ceiling(1584 * layout.RasterDpi / 96d) * Math.Ceiling(2448 * layout.RasterDpi / 96d) * 4 < 64 * 1024 * 1024);
    }
}
