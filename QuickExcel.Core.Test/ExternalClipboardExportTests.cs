using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class ExternalClipboardExportTests
{
    private static WorkbookSnapshot Book(string format, double value, ExcelDateSystem system)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Source", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(value))))]);
        return new(Guid.NewGuid(), 0, system, [sheet], styles: [new() { NumberFormatCode = format }]);
    }

    [Theory]
    [InlineData("m")]
    [InlineData("mm")]
    [InlineData("mmm")]
    [InlineData("mmmm")]
    [InlineData("mmmmm")]
    [InlineData("[ENG]mmmm")]
    [InlineData("[ENG]mmm")]
    [InlineData("[>=1]mmm;0.00")]
    public void MonthOnlyDatesKeepCalendarMeaningAcrossWorkbookDateSystems(string format)
    {
        var book = Book(format, 43538.5, ExcelDateSystem.Mac1904);
        var capture = WorkbookClipboard.Capture(book, book.ActiveSheetId, new(0, 0, 0, 0));
        var source = SpreadsheetClipboardCodec.Capture(book, capture);
        Assert.Equal(new ExcelSerialDate(43538.5, ExcelDateSystem.Mac1904), source.Cells[0].Date);
        Assert.False(CellDisplayFormatter.IsTimeOnlyFormat(book.Styles[0], 43538.5));
        var html = HtmlClipboardCodec.Unwrap(HtmlClipboardCodec.Write(source));
        Assert.Contains("2023-03-15 12:00:00.000", html);
        Assert.DoesNotContain("x:num=", html);
        var portable = SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(source));
        var target = Book("General", 0, ExcelDateSystem.Windows1900);
        using var session = new DocumentSession(target);
        session.Execute(new PasteExternalCellsCommand(portable, target.ActiveSheetId, new(0, 0, 0, 0), 0));
        Assert.Equal(45000.5, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }

    [Theory]
    [InlineData("hh:mm:ss")]
    [InlineData("mm:ss")]
    [InlineData("[h]:mm:ss")]
    [InlineData("[m]:ss")]
    [InlineData("0.000000000")]
    [InlineData("0.00\" months\"")]
    public void TimesDurationsAndNumbersKeepRawPrecisionAcrossInstances(string format)
    {
        const double number = 12.123456789;
        var book = Book(format, number, ExcelDateSystem.Mac1904);
        var capture = WorkbookClipboard.Capture(book, book.ActiveSheetId, new(0, 0, 0, 0));
        var source = SpreadsheetClipboardCodec.Capture(book, capture);
        Assert.Null(source.Cells[0].Date);
        Assert.Contains("x:num=\"12.123456789\"", HtmlClipboardCodec.Unwrap(HtmlClipboardCodec.Write(source)));
        var portable = SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(source));
        Assert.Null(portable.Cells[0].Date);
        Assert.Equal(number, portable.Cells[0].Value.NumberValue);
        var target = Book("General", 0, ExcelDateSystem.Windows1900);
        using var session = new DocumentSession(target);
        session.Execute(new PasteExternalCellsCommand(portable, target.ActiveSheetId, new(0, 0, 0, 0), 0));
        Assert.Equal(number, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExportCannotSilentlyOmitCoveredMergedCellContent(bool formula)
    {
        var hidden = new ExternalClipboardCell(formula ? CellValue.Blank : CellValue.FromText("keep"), new(),
            Formula: formula ? "=1+2" : null, MissingResult: formula);
        var source = new ExternalClipboardRegion(1, 2, [new(CellValue.FromText("anchor"), new()), hidden], [new(0, 0, 0, 1)]);
        Assert.Throws<NotSupportedException>(() => SpreadsheetClipboardCodec.Write(source, formula));
        Assert.Throws<NotSupportedException>(() => HtmlClipboardCodec.Write(source, formula));
    }

    [Fact]
    public void ExplicitResultCopyFlattensMergesAndRetainsCoveredValues()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Source", [
            new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("anchor")))),
            new(new(0, 1), new(CellContent.FromLiteral(CellValue.FromText("keep"))))], mergedCells: ["A1:B1"]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var capture = WorkbookClipboard.Capture(book, sheet.Id, new(0, 0, 0, 1));
        var source = SpreadsheetClipboardCodec.Capture(book, capture, valuesOnly: true);
        Assert.Empty(source.Merges);
        Assert.Equal("keep", SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(source)).Cells[1].Value.TextValue);
        Assert.Contains("keep", HtmlClipboardCodec.Unwrap(HtmlClipboardCodec.Write(source)));
    }

    [Fact]
    public void HtmlExportHonorsSourceLimitationsAndMissingResultGuards()
    {
        var limited = new ExternalClipboardRegion(1, 1, [new(CellValue.FromText("text"), new())], limitations: ["富文本"]);
        Assert.Throws<NotSupportedException>(() => HtmlClipboardCodec.Write(limited));
        var missing = new ExternalClipboardRegion(1, 1, [new(CellValue.Blank, new(), Formula: "=1+2", MissingResult: true)]);
        Assert.Throws<InvalidOperationException>(() => HtmlClipboardCodec.Write(missing));
        Assert.Contains("x:fmla", HtmlClipboardCodec.Unwrap(HtmlClipboardCodec.Write(missing, formulas: true)));
    }
}
