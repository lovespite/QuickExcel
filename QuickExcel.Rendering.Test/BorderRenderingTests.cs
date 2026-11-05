using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class BorderRenderingTests
{
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(3, true)]
    public void MergedPerimeterRetainsGapsAndDifferentStrokes(int side, bool neighbor)
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        var styles = new List<StyleDefinition> { new() };
        foreach (var (index, kind, color) in new[] { (1, 1, "FFFF0000"), (3, 5, "FF0000FF") })
        {
            var edge = neighbor ? side ^ 1 : side;
            styles.Add(edge switch
            {
                0 => new() { LeftBorder = kind, LeftBorderColor = color },
                1 => new() { RightBorder = kind, RightBorderColor = color },
                2 => new() { TopBorder = kind, TopBorderColor = color },
                _ => new() { BottomBorder = kind, BottomBorderColor = color }
            });
            var position = side switch
            {
                0 => new CellPosition(index, neighbor ? 0 : 1),
                1 => new CellPosition(index, neighbor ? 4 : 3),
                2 => new CellPosition(neighbor ? 0 : 1, index),
                _ => new CellPosition(neighbor ? 4 : 3, index)
            };
            cells[position] = new(CellContent.FromLiteral(CellValue.Blank, styles.Count - 1));
        }
        var sheet = new WorksheetSnapshot(SheetId.New(), "Merged", cells, mergedCells: ["B2:D4"], printing: new()
        {
            Margins = new(0, 0, 0, 0, 0, 0), Setup = new() { Scale = 100 },
            Areas = [new(PrintRangeKind.Cells, 0, 0, 4, 4)]
        });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles);
        using var report = new SkiaReport(book, sheet, "merged.xlsx");
        using var bitmap = new SKBitmap(800, 1100); using var canvas = new SKCanvas(bitmap);
        report.DrawPage(canvas, 0);
        SKColor[] EdgePixels(int index)
        {
            var x = side < 2 ? DrawingGeometry.ColumnStart(sheet, side == 0 ? 1 : 4)
                : DrawingGeometry.ColumnStart(sheet, index) + DrawingGeometry.ColumnSize(sheet, index) / 2;
            var y = side >= 2 ? DrawingGeometry.RowStart(sheet, side == 2 ? 1 : 4)
                : DrawingGeometry.RowStart(sheet, index) + DrawingGeometry.RowSize(sheet, index) / 2;
            return Enumerable.Range(-2, 5).Select(offset => bitmap.GetPixel((int)x + (side < 2 ? offset : 0), (int)y + (side >= 2 ? offset : 0))).ToArray();
        }
        Assert.Contains(EdgePixels(1), p => p.Red > 200 && p.Green < 80 && p.Blue < 80);
        Assert.All(EdgePixels(2), p => Assert.Equal(SKColors.White, p));
        Assert.Contains(EdgePixels(3), p => p.Blue > 200 && p.Red < 80 && p.Green < 80);
    }

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
