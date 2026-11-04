using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class BorderRenderingTests
{
    [Fact]
    public void DoubleBottomAndRightRetainTwoSeparatedStrokesAfterNeighborFills()
    {
        using var bitmap = new SKBitmap(100, 100); using var canvas = new SKCanvas(bitmap); using var renderer = new SkiaRenderer(); canvas.Clear(SKColors.White);
        var style = new StyleDefinition { BottomBorder = 6, RightBorder = 6, BottomBorderColor = "FFFF0000", RightBorderColor = "FFFF0000" };
        renderer.DrawCell(canvas, new(new(10, 10, 40, 40), new("", style)), RenderOutput.Screen, borders: false);
        renderer.DrawCell(canvas, new(new(50, 10, 40, 40), new("", new())), RenderOutput.Screen, borders: false);
        renderer.DrawCell(canvas, new(new(10, 50, 40, 40), new("", new())), RenderOutput.Screen, borders: false);
        SkiaRenderer.DrawBorders(canvas, new(10, 10, 40, 40), style);
        // Both strokes survive the background pass with a clean one-pixel gap.
        bool Red(int x, int y) => bitmap.GetPixel(x, y).Red > bitmap.GetPixel(x, y).Green;
        Assert.True(Red(25, 49)); Assert.True(Red(25, 51)); Assert.True(Red(49, 25)); Assert.True(Red(51, 25));
        Assert.Equal(SKColors.White, bitmap.GetPixel(25, 50)); Assert.Equal(SKColors.White, bitmap.GetPixel(50, 25));
    }

    [Fact]
    public void ThickerBordersHaveMoreInkAndDashDotPatternsHaveGaps()
    {
        int Ink(int kind)
        {
            using var bitmap = new SKBitmap(160, 16); using var canvas = new SKCanvas(bitmap); canvas.Clear(SKColors.White);
            SkiaRenderer.DrawBorderLine(canvas, kind, "FF0000FF", new(4, 8), new(156, 8));
            return bitmap.Pixels.Sum(p => 255 - p.Red);
        }
        Assert.True(Ink(7) < Ink(1)); Assert.True(Ink(1) < Ink(2)); Assert.True(Ink(2) < Ink(5));
        foreach (var kind in new[] { 3, 4, 8, 9, 10, 11, 12, 13 }) { Assert.True(Ink(kind) > 10); Assert.True(Ink(kind) < Ink(kind is 3 or 4 or 9 or 11 ? 1 : 2)); }
    }

    [Fact]
    public void DiagonalDirectionsCrossCellInteriorAndMonochromeRemovesColor()
    {
        using var bitmap = new SKBitmap(80, 80); using var canvas = new SKCanvas(bitmap); canvas.Clear(SKColors.White);
        var style = new StyleDefinition { DiagonalBorder = 5, DiagonalBorderColor = "FFFF0000", DiagonalUp = true, DiagonalDown = true };
        SkiaRenderer.DrawBorders(canvas, new(10, 10, 60, 60), style);
        Assert.Equal(SKColors.Red, bitmap.GetPixel(25, 25)); Assert.Equal(SKColors.Red, bitmap.GetPixel(25, 55));
        canvas.Clear(SKColors.White); SkiaRenderer.DrawBorders(canvas, new(10, 10, 60, 60), style, true); Assert.Equal(SKColors.Black, bitmap.GetPixel(25, 25));
    }
}
