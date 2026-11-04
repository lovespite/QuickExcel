using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;
using UglyToad.PdfPig;

namespace QuickExcel.Rendering.Test;

public sealed class ChartRenderingTests
{
    private static (WorkbookSnapshot Book, WorksheetChart Chart) Chart(WorksheetChartKind kind, double[] xs, double[] ys, string[]? categories = null)
    {
        var cells = xs.SelectMany((x, i) => new[] { new KeyValuePair<CellPosition, CellSnapshot>(new(i + 1, 0), new(CellContent.FromLiteral(categories is null ? CellValue.FromNumber(x) : CellValue.FromText(categories[i])))),
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
    [Theory]
    [InlineData(WorksheetChartKind.Column)]
    [InlineData(WorksheetChartKind.Line)]
    [InlineData(WorksheetChartKind.SmoothLine)]
    [InlineData(WorksheetChartKind.Waterfall)]
    public void CategoryLabelsFitAndStayCenteredOnTheirDataPoints(WorksheetChartKind kind)
    {
        var categories = Enumerable.Range(120, 9).Select(i => $"X{i}").ToArray();
        var (book, chart) = Chart(kind, Enumerable.Range(1, 9).Select(i => (double)i).ToArray(), Enumerable.Repeat(2d, 9).ToArray(), categories);
        chart = chart with { Appearance = chart.Appearance with { Axis = new(FontFamily: "Arial") } };
        using var renderer = new SkiaRenderer(); using var stream = new MemoryStream();
        using (var document = SKDocument.CreatePdf(stream))
        { var canvas = document.BeginPage(500, 300); renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart); document.EndPage(); document.Close(); }
        using var pdf = PdfDocument.Open(stream.ToArray());
        var labels = pdf.GetPage(1).GetWords().Where(w => w.Text.StartsWith('X')).ToArray();
        Assert.Equal(categories, labels.Select(w => w.Text));
        for (var i = 0; i < labels.Length; i++)
        {
            var bounds = labels[i].BoundingBox;
            Assert.InRange((bounds.Left + bounds.Right) / 2, 62 + (i + .5) * 418 / 9 - 2, 62 + (i + .5) * 418 / 9 + 2);
            Assert.InRange(bounds.Left, 0, 500); Assert.InRange(bounds.Right, 0, 500);
            Assert.InRange(bounds.Bottom, 0, 40); Assert.InRange(bounds.Top, 0, 40);
        }
    }

    [Theory]
    [InlineData(9, 0, 9)]
    [InlineData(14, 45, 14)]
    [InlineData(18, 90, 18)]
    [InlineData(40, 90, 20)]
    public void CrowdedCategoriesRotateBeforeAnyLabelsAreSkipped(int count, int expectedAngle, int expectedCount)
    {
        var categories = Enumerable.Range(120, count).Select(i => $"X{i}").ToArray();
        var (book, chart) = Chart(WorksheetChartKind.Column, Enumerable.Range(1, count).Select(i => (double)i).ToArray(), Enumerable.Repeat(2d, count).ToArray(), categories);
        chart = chart with { Appearance = chart.Appearance with { Axis = new(FontFamily: "Arial") } };
        using var renderer = new SkiaRenderer(); using var stream = new MemoryStream();
        using (var document = SKDocument.CreatePdf(stream))
        { var canvas = document.BeginPage(500, 300); renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart); document.EndPage(); document.Close(); }
        using var pdf = PdfDocument.Open(stream.ToArray());
        var letters = pdf.GetPage(1).Letters;
        var starts = Enumerable.Range(0, letters.Count).Where(i => letters[i].Value == "X").ToArray();
        Assert.Equal(expectedCount, starts.Length);
        var stride = count / expectedCount;
        for (var i = 0; i < starts.Length; i++)
        {
            var label = letters.Skip(starts[i]).Take(4).ToArray();
            Assert.Equal(categories[i * stride], string.Concat(label.Select(l => l.Value)));
            var from = label[0].StartBaseLine; var to = label[^1].EndBaseLine;
            var angle = Math.Atan2(to.Y - from.Y, to.X - from.X) * 180 / Math.PI;
            Assert.InRange(angle, expectedAngle - .1, expectedAngle + .1);
            foreach (var letter in label)
            {
                Assert.InRange(letter.BoundingBox.Left, 0, 500); Assert.InRange(letter.BoundingBox.Right, 0, 500);
                Assert.InRange(letter.BoundingBox.Bottom, 0, 100); Assert.InRange(letter.BoundingBox.Top, 0, 100);
            }
            if (expectedAngle == 90)
            {
                var center = label.Average(l => l.BoundingBox.Centroid.X);
                Assert.InRange(center, 62 + (i * stride + .5) * 418 / count - 3, 62 + (i * stride + .5) * 418 / count + 3);
            }
        }
    }

    [Theory]
    [InlineData(6, 9, "-East-region", 45)]
    [InlineData(12, 18, "", 90)]
    public void RotatedLabelsReserveRoomForCompleteTextAndBottomLegend(int count, double fontSize, string suffix, int expectedAngle)
    {
        var categories = Enumerable.Range(120, count).Select(i => $"X{i}{suffix}").ToArray();
        var (book, chart) = Chart(WorksheetChartKind.Column, Enumerable.Range(1, count).Select(i => (double)i).ToArray(), Enumerable.Repeat(2d, count).ToArray(), categories);
        chart = chart with { Legend = "b", Series = [chart.Series[0] with { Name = "Series", NameReference = null }],
            Appearance = chart.Appearance with { Axis = new(FontFamily: "Arial", Size: fontSize) } };
        using var renderer = new SkiaRenderer(); using var stream = new MemoryStream();
        using (var document = SKDocument.CreatePdf(stream))
        { var canvas = document.BeginPage(500, 300); renderer.DrawChart(canvas, new(0, 0, 500, 300), book, chart); document.EndPage(); document.Close(); }
        using var pdf = PdfDocument.Open(stream.ToArray());
        var letters = pdf.GetPage(1).Letters;
        var starts = Enumerable.Range(0, letters.Count).Where(i => letters[i].Value == "X").ToArray();
        Assert.Equal(count, starts.Length);
        for (var i = 0; i < starts.Length; i++)
        {
            var label = letters.Skip(starts[i]).Take(categories[i].Length).ToArray();
            Assert.Equal(categories[i], string.Concat(label.Select(l => l.Value)));
            var from = label[0].StartBaseLine; var to = label[^1].EndBaseLine;
            Assert.InRange(Math.Atan2(to.Y - from.Y, to.X - from.X) * 180 / Math.PI, expectedAngle - .1, expectedAngle + .1);
            foreach (var letter in label)
            {
                Assert.InRange(letter.BoundingBox.Left, 0, 500); Assert.InRange(letter.BoundingBox.Right, 0, 500);
                Assert.InRange(letter.BoundingBox.Bottom, 32, 258); Assert.InRange(letter.BoundingBox.Top, 32, 258);
            }
        }
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
