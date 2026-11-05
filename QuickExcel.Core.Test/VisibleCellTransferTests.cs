using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Calculation;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class VisibleCellTransferTests
{
    private static WorkbookSnapshot Fixture()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Source", [
            new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("00123"), 1))),
            new(new(1, 0), N(999)), new(new(2, 0), new(CellContent.FromLiteral(CellValue.FromText("=1+1"), 1))),
            new(new(0, 1), new(CellContent.FromFormula("=A1+1"), new(CellValue.FromNumber(124), CalculationState.Fresh, 0, 0, 0))),
            new(new(2, 1), new(CellContent.FromFormula("=A3+2"), new(CellValue.FromNumber(3), CalculationState.Fresh, 0, 0, 0)))
        ], rows: [new(1, new(Hidden: true))]);
        var target = Sheet("Target", (5, 2, T("keep")), (5, 3, T("also keep")));
        target = new(target.Id, target.Name, target.Cells, rows: [new(5, new(Hidden: true))]);
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source, target], styles: [new(), new() { Bold = true, NumberFormatCode = "@" }]);
    }
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void TypedAndFormulaRowsMapAcrossGapsAndSheetsWithOneUndo(bool formulas)
    {
        var book = Fixture(); var capture = VisibleCellTransfer.Capture(book, book.Sheets[0].Id, new(0, 0, 2, 1), formulas);
        var source = SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(capture, formulas));
        Assert.Equal(2, source.Rows); Assert.Equal(CellValue.FromText("00123"), source.Cells[0].Value);
        using var session = new DocumentSession(book); var id = book.Sheets[1].Id;
        session.Execute(new PasteVisibleCellsCommand(source, id, new(4, 2, 6, 3), 0, formulas ? CellPasteMode.Formulas : CellPasteMode.All));
        var result = session.Snapshot.GetSheet(id);
        Assert.Equal("00123", result.GetCell(4, 2)!.Value.TextValue); Assert.Equal("=1+1", result.GetCell(6, 2)!.Value.TextValue);
        Assert.Equal("keep", result.GetCell(5, 2)!.Value.TextValue); Assert.Equal("also keep", result.GetCell(5, 3)!.Value.TextValue);
        if (formulas)
        { Assert.Equal("=(C5+1)", result.GetCell(4, 3)!.Content.FormulaText); Assert.Equal("=(C7+2)", result.GetCell(6, 3)!.Content.FormulaText); }
        else Assert.True(session.Snapshot.Styles[result.GetStyleId(4, 2)].Bold);
        Assert.True(session.Undo()); Assert.Equal(book.Sheets[1].Cells, session.Snapshot.GetSheet(id).Cells); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); Assert.All(result.Cells, pair => Assert.Equal(pair.Value.Content, session.Snapshot.GetSheet(id).Cells[pair.Key].Content));
    }
    [Fact]
    public void FilteringAndManualHiddenRowsBothContributeToOrderedMapping()
    {
        var sheet = Sheet("S", (0, 0, T("H")), (1, 0, T("A")), (2, 0, T("B")), (3, 0, T("A")), (4, 0, T("A")));
        sheet = new(sheet.Id, sheet.Name, sheet.Cells, rows: [new(3, new(Hidden: true))], data: new(new(new(0, 0, 4, 0), [new(0, new(values: ["A"]))])));
        Assert.Equal(new[] { 1, 4 }, VisibleCellTransfer.Rows(Book(sheet), sheet.Id, new(1, 0, 4, 0)));
    }
    [Theory]
    [InlineData("count")] [InlineData("revision")] [InlineData("protected")] [InlineData("validation")] [InlineData("rich")] [InlineData("merge")]
    public void EveryFailureRollsBackAllRowsAndNewStyles(string kind)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Target", kind == "rich" ? [new(new(2, 0), new(CellContent.FromRichText(new([new("rich")]))))] : null,
            rows: [new(1, new(Hidden: true))], mergedCells: kind == "merge" ? ["A3:B3"] : null,
            data: kind == "protected" ? new() { Protection = new() } : kind == "validation" ? new() { Validations = [new([new(2, 0, 2, 0)], ValidationKind.Whole, "=1", "=5")] } : null);
        var book = Book(sheet); using var session = new DocumentSession(book, validation: new WorkbookValidation());
        var source = new ExternalClipboardRegion(2, 1, [new(CellValue.FromNumber(2), new() { Bold = true }), new(CellValue.FromNumber(9), new() { Italic = true })]);
        Assert.ThrowsAny<Exception>(() => session.Execute(new PasteVisibleCellsCommand(source, sheet.Id, new(0, 0, kind == "count" ? 3 : 2, 0), kind == "revision" ? 1 : 0)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }
    [Fact]
    public void CaptureRejectsOldResultsUnknownAttachmentsHiddenColumnsAndUnsupportedFormulas()
    {
        var stale = Sheet("S", (0, 0, F("=1")));
        Assert.Throws<InvalidOperationException>(() => VisibleCellTransfer.Capture(Book(stale), stale.Id, new(0, 0, 0, 0)));
        var book = Fixture(); var first = book.Sheets[0];
        var hidden = new WorksheetSnapshot(first.Id, first.Name, columns: [new(0, new(Hidden: true))]);
        Assert.Throws<NotSupportedException>(() => VisibleCellTransfer.Rows(Book(hidden), hidden.Id, new(0, 0, 1, 1)));
        var rich = Sheet("S", (0, 0, new(CellContent.FromRichText(new([new("rich")])))));
        Assert.Throws<NotSupportedException>(() => VisibleCellTransfer.Capture(Book(rich), rich.Id, new(0, 0, 0, 0)));
        var absolute = Sheet("S", (0, 0, new(CellContent.FromFormula("=$B$2"), new(CellValue.FromNumber(1), CalculationState.Fresh, 0, 0, 0))));
        Assert.Throws<NotSupportedException>(() => VisibleCellTransfer.Capture(Book(absolute), absolute.Id, new(0, 0, 0, 0), true));
    }
    [Fact]
    public void DateValuesConvertEpochAndPersistBothFormats()
    {
        var source = new ExternalClipboardRegion(1, 1, [new(CellValue.FromNumber(45000), new() { NumberFormatCode = "yyyy-mm-dd" }, Date: new(45000, ExcelDateSystem.Windows1900))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904, [new(SheetId.New(), "S")]); using var session = new DocumentSession(book);
        session.Execute(new PasteVisibleCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), 0));
        Assert.Equal(43538, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "visible-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        foreach (var ext in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "visible." + ext); var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            Assert.Equal(43538, reopened.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
            Assert.Equal("yyyy-mm-dd", reopened.Styles[reopened.Sheets[0].GetStyleId(0, 0)].NumberFormatCode);
        }
    }
    [Fact]
    public void ScanBudgetAndCancellationAreBounded()
    {
        var book = Fixture(); var id = book.Sheets[0].Id;
        Assert.Throws<InvalidOperationException>(() => VisibleCellTransfer.Rows(book, id, new(0, 0, 30000, 0)));
        Assert.Throws<OperationCanceledException>(() => VisibleCellTransfer.Capture(book, id, new(0, 0, 2, 0), cancellationToken: new(true)));
    }
}
