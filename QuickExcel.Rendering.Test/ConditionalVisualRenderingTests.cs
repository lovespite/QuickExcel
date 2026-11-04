using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class ConditionalVisualRenderingTests
{
    [Fact]
    public void BarsRenderBehindTextWithAxisAndRespectHideValue()
    {
        using var bitmap = new SKBitmap(120, 30); using var canvas = new SKCanvas(bitmap); using var renderer = new SkiaRenderer();
        var bar = new ConditionalBarVisual(.5, .1, .5, "FFFF0000", false, null, "FF000000");
        var input = new CellRenderInput(new(0, 0, 120, 30), new("123", new() { Horizontal = 3, Vertical = 1 }), Conditional: new(bar, HideValue: true));
        renderer.DrawCell(canvas, input);
        Assert.Equal(SKColors.Red, bitmap.GetPixel(30, 10)); Assert.Equal(SKColors.White, bitmap.GetPixel(90, 10));
        Assert.NotEqual(SKColors.White, bitmap.GetPixel(60, 10));
        renderer.DrawCell(canvas, input with { Conditional = input.Conditional! with { HideValue = false } });
        Assert.Contains(Enumerable.Range(80, 38).SelectMany(x => Enumerable.Range(4, 22).Select(y => bitmap.GetPixel(x, y))), c => c.Red < 100 && c.Green < 100 && c.Blue < 100);
    }

    [Fact]
    public void EveryIconFamilyPaintsAndIsClippedToTheCell()
    {
        using var bitmap = new SKBitmap(32, 32); using var canvas = new SKCanvas(bitmap); using var renderer = new SkiaRenderer();
        foreach (var name in ConditionalIconSet.Names)
            for (var i = 0; i < name[0] - '0'; i++)
            {
                canvas.Clear(SKColors.Magenta);
                renderer.DrawCell(canvas, new(new(4, 4, 20, 20), new("", new()), Conditional: new(Icon: new(name, i))));
                Assert.Equal(SKColors.Magenta, bitmap.GetPixel(2, 2)); Assert.Equal(SKColors.Magenta, bitmap.GetPixel(27, 15));
                Assert.True(Enumerable.Range(5, 18).SelectMany(x => Enumerable.Range(5, 18).Select(y => bitmap.GetPixel(x, y))).Any(c => c != SKColors.White), name + "/" + i);
            }
    }
}
