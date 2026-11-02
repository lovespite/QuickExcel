using QuickExcel.Core;
using QuickExcel.Domain;
using SkiaSharp;

namespace QuickExcel.Rendering.Smoke;

internal static class ReportFixture
{
    internal static WorkbookSnapshot Create()
    {
        var style = new StyleDefinition { FontName = "Noto Sans", EastAsianFontName = "Noto Sans CJK SC", FontSize = 11, Vertical = 1 };
        StyleDefinition[] styles = [style, style with { FontSize = 18, Bold = true, FillColor = "FFE8EFF8", Horizontal = 2 },
            style with { WrapText = true, Vertical = 0 }, style with { TextRotation = 45, FontColor = "FF2F5597" }, style with { ShrinkToFit = true },
            style with { NumberFormatCode = "#,##0.00", BottomBorder = 6, BottomBorderColor = "FF4472C4" }];
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        void Text(int r, int c, string value, int s = 0) => cells[new(r, c)] = new(CellContent.FromLiteral(CellValue.FromText(value), s));
        Text(0, 0, "QuickExcel 中文报表 / Shared rendering", 1); Text(1, 0, "记录 Record"); Text(1, 1, "收入 Revenue");
        for (var r = 2; r < 65; r++) { Text(r, 0, $"Record {r:00}"); cells[new(r, 1)] = new(CellContent.FromLiteral(CellValue.FromNumber(r * 12.5), 5)); }
        cells[new(2, 2)] = new(CellContent.FromRichText(new([new("富文本 "), new("Bold", new(Bold: true, Color: "FF4472C4")), new(" + underline", new(Underline: "double", Strike: true))])));
        Text(4, 2, "统一测量与自动换行。This paragraph wraps without splitting combining marks: e\u0301, A\u030a and supplementary 𠮷.", 2);
        Text(8, 2, "旋转 45°", 3); Text(10, 2, "Shrink to fit 收缩填充", 4);
        Text(24, 2, "Page two / 第二页"); Text(44, 2, "Page three / 第三页");
        using var bitmap = new SKBitmap(80, 50); using (var canvas = new SKCanvas(bitmap)) { canvas.Clear(SKColors.Transparent); using var p = new SKPaint { Color = SKColors.CornflowerBlue }; canvas.DrawCircle(25, 25, 22, p); p.Color = SKColors.Orange; canvas.DrawRect(45, 8, 30, 35, p); }
        using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100);
        var image = new SheetImage(encoded.ToArray(), "png", 8, 3, anchor: new() { Kind = DrawingAnchorKind.OneCell, Row = 8, Column = 3, Width = 120 * 9525, Height = 75 * 9525 });
        var sheet = new WorksheetSnapshot(SheetId.New(), "Report", cells, rows: [new(0, new(42)), new(4, new(42)), new(5, new(42)), new(6, new(Hidden: true)), new(8, new(72))],
            columns: Enumerable.Range(0, 5).Select(c => new KeyValuePair<int, ColumnMetadata>(c, new(Width: c == 0 ? 18 : 14))),
            defaultRowHeight: 28, mergedCells: ["A1:E1", "C3:E3", "C5:E7"], images: [image],
            printing: new() { Areas = [new(PrintRangeKind.Cells, 0, 0, 64, 4)], Titles = [new(PrintRangeKind.Rows, 0, 0, 1, PrintRange.MaxColumns - 1)],
                RowBreaks = [new(24, 0, WorkbookLimits.MaxColumns - 1, true), new(44, 0, WorkbookLimits.MaxColumns - 1, true)],
                Setup = new() { PaperSize = 9, Scale = 95 }, Options = new() { GridLines = true },
                HeaderFooter = new() { OddHeader = "&LQuickExcel&C中文报告&R&F", OddFooter = "&L2026-11-02&C第 &P / &N 页" } });
        WorksheetChart Chart(WorksheetChartKind kind, int row, string title) => new() { Kind = kind, Title = title, Legend = "b",
            Anchor = new() { Kind = DrawingAnchorKind.OneCell, Row = row, Column = 2, Width = 300 * 9525, Height = 230 * 9525 },
            Series = [new("收入 Revenue", null, "Report!A3:A7", "Report!B3:B7") { PointColors = ["FF4472C4", "FFED7D31", "FFA5A5A5", "FFFFC000", "FF70AD47"] }] };
        sheet = sheet.WithData(sheet.Data with { Charts = [Chart(WorksheetChartKind.Column, 12, "收入 / Column"), Chart(WorksheetChartKind.Line, 27, "趋势 / Line"), Chart(WorksheetChartKind.Pie, 47, "占比 / Pie")] });
        return new(Guid.Parse("4dfb655d-8714-4a54-b177-284fa1f73cb2"), 0, ExcelDateSystem.Windows1900, [sheet], "", styles);
    }
}
