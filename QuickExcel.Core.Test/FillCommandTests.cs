using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class FillCommandTests
{
    [Theory]
    [InlineData(CellSeriesKind.Numbers, ExcelDateSystem.Windows1900, 0.1, 0.1, 2, 0.30000000000000004)]
    [InlineData(CellSeriesKind.Numbers, ExcelDateSystem.Windows1900, 3, -2, 7, -11)]
    [InlineData(CellSeriesKind.Days, ExcelDateSystem.Windows1900, 59, 1, 1, 60)]
    [InlineData(CellSeriesKind.Days, ExcelDateSystem.Windows1900, 59, 1, 2, 61)]
    [InlineData(CellSeriesKind.Days, ExcelDateSystem.Mac1904, 0, 1, 7, 7)]
    [InlineData(CellSeriesKind.Months, ExcelDateSystem.Windows1900, 31, 1, 1, 60)]
    [InlineData(CellSeriesKind.Months, ExcelDateSystem.Windows1900, 31, 1, 2, 91)]
    [InlineData(CellSeriesKind.Months, ExcelDateSystem.Windows1900, 60, 1, 1, 89)]
    [InlineData(CellSeriesKind.Months, ExcelDateSystem.Windows1900, 45322.5, 1, 1, 45351.5)]
    [InlineData(CellSeriesKind.Months, ExcelDateSystem.Windows1900, 45322, -1, 3, 45230)]
    [InlineData(CellSeriesKind.Years, ExcelDateSystem.Windows1900, 45351, 1, 1, 45716)]
    [InlineData(CellSeriesKind.Years, ExcelDateSystem.Windows1900, 45351, 1, 4, 46812)]
    [InlineData(CellSeriesKind.Years, ExcelDateSystem.Mac1904, 59, 1, 4, 1520)]
    public void SeriesMatchesIndependentExcelValues(CellSeriesKind kind, ExcelDateSystem system, double start, double step, int index, double expected) =>
        Assert.Equal(expected, FillSeriesCommand.SeriesValue(start, step, index, kind, system));

    [Theory]
    [InlineData("XX1", 1, 1, "XX2")]
    [InlineData("XX1", 1, 9, "XX10")]
    [InlineData("XX1", 2, 2, "XX5")]
    [InlineData("A001", 1, 1, "A002")]
    [InlineData("A001", 1, 99, "A100")]
    [InlineData("001", 1, 1, "002")]
    [InlineData("编号9", -1, 1, "编号8")]
    public void TextNumberSeriesKeepsPrefixAndLeadingZeroWidth(string text, double step, int index, string expected) =>
        Assert.Equal(expected, FillSeriesCommand.TextSeriesValue(text, step, index));

    [Fact]
    public void TextNumberSeriesFillsSelectedCellsWithTheSourceStyleAndOneUndo()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "TextSeries", [new(new(1, 1), new(CellContent.FromLiteral(CellValue.FromText("XX1"), 2)))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }, new() { Italic = true }]);
        using var session = new DocumentSession(book);
        session.Execute(new FillSeriesCommand(sheet.Id, new(1, 1, 1, 4), CellSeriesKind.TextNumber, 1));
        var result = session.Snapshot.Sheets[0];
        Assert.Equal("XX1", result.GetCell(1, 1)!.Value.TextValue);
        Assert.Equal("XX2", result.GetCell(1, 2)!.Value.TextValue);
        Assert.Equal("XX4", result.GetCell(1, 4)!.Value.TextValue);
        Assert.All(Enumerable.Range(1, 4), c => Assert.Equal(2, result.GetStyleId(1, c)));
        session.Undo(); Assert.Equal("XX1", session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.TextValue);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 4));
        session.Redo(); Assert.Equal("XX4", session.Snapshot.Sheets[0].GetCell(1, 4)!.Value.TextValue);
    }

    [Fact]
    public void TextNumberSeriesRejectsMissingDigitsNegativeResultsAndFractionalSteps()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "TextSeries", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("XX")))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromText("X1"))))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]); using var session = new DocumentSession(book);
        Assert.Throws<ArgumentException>(() => session.Execute(new FillSeriesCommand(sheet.Id, new(0, 0, 3, 0), CellSeriesKind.TextNumber, 1)));
        Assert.Throws<ArgumentException>(() => session.Execute(new FillSeriesCommand(sheet.Id, new(1, 0, 3, 0), CellSeriesKind.TextNumber, -5)));
        Assert.Throws<ArgumentException>(() => session.Execute(new FillSeriesCommand(sheet.Id, new(1, 0, 3, 0), CellSeriesKind.TextNumber, .5)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    private static WorkbookSnapshot Book()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Fill", [new(new(2, 2), new(CellContent.FromLiteral(CellValue.FromText("00123"), 1))),
            new(new(3, 2), new(CellContent.FromLiteral(CellValue.FromBoolean(true), 2))),
            new(new(2, 3), new(CellContent.FromFormula("=C3+$A$1", 1))), new(new(3, 3), new(CellContent.FromFormula("=C4+$A$1", 2)))]);
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new() { FontName = "Arial", FontSize = 10 }, new() { Bold = true }, new() { Italic = true }]);
    }

    [Theory]
    [InlineData(2, 2, 7, 3, 6, 3, "=(C7+$A$1)", 1)]
    [InlineData(0, 2, 3, 3, 1, 3, "=(C2+$A$1)", 2)]
    [InlineData(2, 2, 3, 7, 2, 7, "=(G3+$A$1)", 1)]
    [InlineData(2, 0, 3, 3, 2, 1, "=(A3+$A$1)", 1)]
    public void FillRepeatsTypedBlocksInAllDirectionsWithAtomicUndo(int top, int left, int bottom, int right, int row, int column, string formula, int style)
    {
        var book = Book(); var sheet = book.Sheets[0]; using var session = new DocumentSession(book);
        session.Execute(new FillCellsCommand(sheet.Id, new(2, 2, 3, 3), new(top, left, bottom, right)));
        var result = session.Snapshot.Sheets[0]; var cell = result.GetCell(row, column)!;
        Assert.Equal(formula, cell.Content.FormulaText); Assert.Equal(style, cell.Content.StyleId);
        Assert.Equal("00123", result.GetCell(row - (style == 2 ? 1 : 0), column - 1)!.Value.TextValue);
        Assert.Equal(sheet.GetCell(2, 3)!.Content, result.GetCell(2, 3)!.Content);
        session.Undo(); Assert.False(session.CanUndo); EqualContent(sheet, session.Snapshot.Sheets[0]);
        session.Redo(); EqualContent(result, session.Snapshot.Sheets[0]);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void SeriesAndClearFormatsPreserveTypesAndSaveReopen(string extension)
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(59), 1))]));
        session.Execute(new FillSeriesCommand(id, new(0, 0, 4, 0), CellSeriesKind.Days, 1));
        Assert.Equal(60, session.Snapshot.Sheets[0].GetCell(1, 0)!.Value.NumberValue);
        Assert.Equal(1, session.Snapshot.Sheets[0].GetCell(4, 0)!.Content.StyleId);
        session.Execute(new ClearCellFormatsCommand([new(id, 1, 0), new(id, 2, 3)]));
        Assert.Equal(0, session.Snapshot.Sheets[0].GetCell(1, 0)!.Content.StyleId);
        Assert.Equal(60, session.Snapshot.Sheets[0].GetCell(1, 0)!.Value.NumberValue);
        Assert.Equal("=C3+$A$1", session.Snapshot.Sheets[0].GetCell(2, 3)!.Content.FormulaText);
        session.Undo(); Assert.Equal(1, session.Snapshot.Sheets[0].GetCell(1, 0)!.Content.StyleId); session.Redo();
        var path = Path.Combine(AppContext.BaseDirectory, "artifacts", "fill-" + Guid.NewGuid().ToString("N") + "." + extension);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!); var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot);
        EqualContent(session.Snapshot.Sheets[0], storage.Open(path).Sheets[0]);
    }

    [Fact]
    public void ClearContentsAndFormatsHaveSeparateUndoAndRichTextContracts()
    {
        var rich = new RichCellText([new("A", new(Color: "FFFF0000")), new("B", new(Bold: true))]); var id = SheetId.New();
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(id, "S", [new(new(0, 0), new(CellContent.FromRichText(rich, 1)))],
            annotations: new(notes: [new(new(0, 0), new("author", new([new("note")])))]))], styles: [new() { FontName = "Arial" }, new() { Bold = true }]);
        using var session = new DocumentSession(book);
        var result = session.Execute(new ClearCellFormatsCommand([new(id, 0, 0)])); Assert.True(result.CalculationUnchanged);
        var cell = session.Snapshot.Sheets[0].GetCell(0, 0)!; Assert.Equal("AB", cell.Value.TextValue); Assert.Equal(0, cell.Content.StyleId);
        Assert.Equal(rich, cell.Content.RichText); Assert.Single(session.Snapshot.Sheets[0].Annotations.Notes);
        session.Undo(); Assert.Equal(rich, session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
        session.Execute(new ClearCellsCommand([new(id, 0, 0)])); Assert.Equal(CellValue.Blank, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(1, session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.StyleId); Assert.Single(session.Snapshot.Sheets[0].Annotations.Notes);
    }

    [Fact]
    public void FillAndSeriesRejectAmbiguousOrInvalidRangesWithoutPartialWrites()
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        Assert.Throws<ArgumentException>(() => session.Execute(new FillCellsCommand(id, new(2, 2, 3, 3), new(2, 2, 5, 5))));
        Assert.Throws<ArgumentException>(() => session.Execute(new FillSeriesCommand(id, new(2, 2, 8, 2), CellSeriesKind.Numbers, 1)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
        session.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 4, 1, hidden: true)); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new FillCellsCommand(id, new(2, 2, 3, 3), new(2, 2, 7, 3)))); Assert.Same(before, session.Snapshot);
        session.Execute(new SetCellsCommand([new(new(id, 8, 0), CellContent.FromLiteral(CellValue.FromNumber(2958465)))])); before = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(new FillSeriesCommand(id, new(8, 0, 10, 0), CellSeriesKind.Days, 1))); Assert.Same(before, session.Snapshot);
        Assert.Throws<ArgumentException>(() => session.Execute(new FillSeriesCommand(id, new(8, 0, 10, 0), CellSeriesKind.Months, .5))); Assert.Same(before, session.Snapshot);
        Assert.Throws<ArgumentException>(() => new ClearCellFormatsCommand(Enumerable.Repeat(new CellAddress(id, 0, 0), 20001)));
    }
    private static void EqualContent(WorksheetSnapshot expected, WorksheetSnapshot actual) =>
        Assert.Equal(expected.Cells.ToDictionary(p => p.Key, p => p.Value.Content), actual.Cells.ToDictionary(p => p.Key, p => p.Value.Content));
}
