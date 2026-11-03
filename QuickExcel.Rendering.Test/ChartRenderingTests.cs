using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class ChartRenderingTests
{
    private static (WorkbookSnapshot Book, WorksheetChart Chart) Chart(WorksheetChartKind kind, double[] xs, double[] ys)
    {
        var cells = xs.SelectMany((x, i) => new[] { new KeyValuePair<CellPosition, CellSnapshot>(new(i + 1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(x)))),
            new KeyValuePair<CellPosition, CellSnapshot>(new(i + 1, 1), new(CellContent.FromLiteral(CellValue.FromNumber(ys[i])))) });
        var sheet = new WorksheetSnapshot(SheetId.New(), "Plot", cells);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, xs.Length, 1), kind, DrawingGeometry.FromBounds(sheet, new(0, 0, 500, 300), DrawingAnchorKind.OneCell)) with { Legend = null };
        return (new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]), chart);
    }
    private static SKBitmap Render(WorkbookSnapshot book, WorksheetChart chart)
    {
        using var renderer = new SkiaRenderer(); var image = new SKBitmap(500, 300); using var canvas = new SKCanvas(image);
        canvas.Clear(SKColors.White); renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart); return image;
    }
    [Fact]
    public void ScatterPlacesNumericXValuesProportionallyRatherThanAsCategories()
    {
        var (book, chart) = Chart(WorksheetChartKind.Scatter, [1, 2, 100], [2, 2, 2]); using var image = Render(book, chart);
        int Blue(int left, int right) => Enumerable.Range(left, right - left).Sum(x => Enumerable.Range(42, 12).Count(y => image.GetPixel(x, y) is var c && c.Blue > c.Red + 80 && c.Blue > c.Green + 40));
        Assert.True(Blue(62, 74) > 12); Assert.Equal(0, Blue(150, 400)); Assert.True(Blue(475, 481) > 4);
    }
    [Fact]
    public void SmoothLineUsesCurves()
    {
        var (book, line) = Chart(WorksheetChartKind.Line, [1, 2, 3, 4], [1, 5, 2, 4]);
        using var straight = Render(book, line); using var smooth = Render(book, line with { Kind = WorksheetChartKind.SmoothLine });
        var differences = Enumerable.Range(62, 418).Sum(x => Enumerable.Range(42, 218).Count(y => straight.GetPixel(x, y) != smooth.GetPixel(x, y)));
        Assert.True(differences > 100);
    }
    [Fact]
    public void WaterfallPaintsNegativeContributionsAndTotalsWithTheirOwnColors()
    {
        var (book, chart) = Chart(WorksheetChartKind.Waterfall, [1, 2, 3], [100, -30, 70]);
        chart = chart with { Series = [chart.Series[0] with { TotalPoints = [2], ShowValues = false, PointColors = ["FFFF00FF", "", ""] }] };
        using var image = Render(book, chart);
        Assert.Equal(new SKColor(237, 125, 49), image.GetPixel(271, 80));
        Assert.Equal(new SKColor(165, 165, 165), image.GetPixel(410, 200));
    }
    [Fact]
    public void ImageRotationUsesItsCenterAndRetainsPixelsOutsideOriginalRectangle()
    {
        using var source = new SKBitmap(80, 20); source.Erase(SKColors.Magenta);
        using var output = new SKBitmap(220, 220); using var canvas = new SKCanvas(output); canvas.Clear(SKColors.White);
        SkiaRenderer.DrawImage(canvas, source, SKRect.Create(100, 100, 80, 20), false, 90);
        Assert.Equal(SKColors.Magenta, output.GetPixel(140, 75)); Assert.Equal(SKColors.White, output.GetPixel(105, 110));
    }
}
