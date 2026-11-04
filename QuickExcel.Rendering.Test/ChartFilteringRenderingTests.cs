using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class ChartFilteringRenderingTests
{
    [Fact]
    public void HidingFirstSeriesKeepsRemainingSeriesMarkerAndLineOptions()
    {
        var cells = Enumerable.Range(1, 2).SelectMany(i => Enumerable.Range(0, 3).Select(c => new KeyValuePair<CellPosition, CellSnapshot>(new(i, c), new(CellContent.FromLiteral(CellValue.FromNumber(c == 0 ? i : 2))))));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Plot", cells); var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 2, 2), WorksheetChartKind.Scatter, new() { Kind = DrawingAnchorKind.Absolute, Width = 4762500, Height = 2857500 }) with { Legend = null };
        chart = chart with { Series = [chart.Series[0] with { IsFiltered = true, ConnectPoints = true, LineWidth = 20 }, chart.Series[1] with { Color = "FF00AA00", ConnectPoints = false, Marker = "circle" }] };
        using var renderer = new SkiaRenderer(); using var bitmap = new SKBitmap(500, 300); using var canvas = new SKCanvas(bitmap);
        renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart);
        int Green(int left, int right) => Enumerable.Range(left, right - left).Sum(x => Enumerable.Range(40, 20).Count(y => bitmap.GetPixel(x, y) is var c && c.Green > c.Red + 40 && c.Green > c.Blue + 40));
        Assert.True(Green(60, 75) > 0); Assert.Equal(0, Green(150, 400));
    }
    [Theory]
    [InlineData(WorksheetChartKind.Column)] [InlineData(WorksheetChartKind.Line)] [InlineData(WorksheetChartKind.SmoothLine)]
    [InlineData(WorksheetChartKind.Pie)] [InlineData(WorksheetChartKind.Scatter)] [InlineData(WorksheetChartKind.Waterfall)] [InlineData(WorksheetChartKind.Funnel)]
    public void EmptyFilterSelectionRendersForEveryChartKind(WorksheetChartKind kind)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Plot"); var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 2, 1), kind, new() { Kind = DrawingAnchorKind.Absolute, Width = 4762500, Height = 2857500 });
        using var renderer = new SkiaRenderer(); using var bitmap = new SKBitmap(500, 300); using var canvas = new SKCanvas(bitmap);
        renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart with { Series = [chart.Series[0] with { IsFiltered = true }] });
        renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart with { Series = [chart.Series[0] with { HiddenPoints = [0, 1] }] });
        Assert.Equal(SKColors.White, bitmap.GetPixel(250, 290));
    }
}
