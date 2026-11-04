using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class AxisClipboardTests
{
    [Theory]
    [InlineData(SheetAxis.Row, 0)]
    [InlineData(SheetAxis.Row, 4)]
    [InlineData(SheetAxis.Column, 0)]
    [InlineData(SheetAxis.Column, 4)]
    public void InsertCopiesContentDimensionsAndShiftedReferencesInOneHistoryEntry(SheetAxis axis, int target)
    {
        var row = axis == SheetAxis.Row;
        var sheet = new WorksheetSnapshot(SheetId.New(), "S", [
            new(new(1, 1), new(CellContent.FromLiteral(CellValue.FromNumber(42), 1))),
            new(row ? new(1, 2) : new(2, 1), new(CellContent.FromFormula("=$B$2+B2", 1)))],
            rows: [new(1, new(Height: 37))], columns: [new(1, new(Width: 23))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book);
        var copy = AxisClipboard.Capture(book, sheet.Id, axis, 1, 1);
        session.Execute(new InsertCopiedAxisCommand(copy, sheet.Id, target));
        var result = session.Snapshot.GetSheet(sheet.Id);
        Assert.Equal(CellValue.FromNumber(42), result.GetCell(row ? target : 1, row ? 1 : target)!.Value);
        var formula = result.GetCell(row ? target : 2, row ? 2 : target)!.Content.FormulaText;
        var absolute = target == 0 ? row ? "$B$3" : "$C$2" : "$B$2";
        var relative = row ? "B" + (target + 1) : (target == 0 ? "A" : "E") + "2";
        Assert.Contains(absolute, formula); Assert.Contains("+" + relative, formula);
        Assert.Equal(1, result.GetStyleId(row ? target : 1, row ? 1 : target));
        if (row) Assert.Equal(37, result.Rows[target].Height); else Assert.Equal(23, result.Columns[target].Width);
        Assert.Equal(1, session.Revision);
        Assert.True(session.Undo()); Assert.Equal(sheet.Cells.ToDictionary(p => p.Key, p => p.Value.Content), session.Snapshot.GetSheet(sheet.Id).Cells.ToDictionary(p => p.Key, p => p.Value.Content)); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); Assert.Equal(result.Cells.ToDictionary(p => p.Key, p => p.Value.Content), session.Snapshot.GetSheet(sheet.Id).Cells.ToDictionary(p => p.Key, p => p.Value.Content));
    }

    [Fact]
    public void CopiesMultipleRowsWithMergesAndClearsInheritedNeighborFormatting()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "S", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("标题"), 1))),
            new(new(3, 9), new(CellContent.FromLiteral(CellValue.Blank, 1)))], mergedCells: ["A1:C1"]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book);
        var source = AxisClipboard.Capture(book, sheet.Id, SheetAxis.Row, 0, 2);
        Assert.Equal(3, source.Cells.Columns);
        session.Execute(new InsertCopiedAxisCommand(source, sheet.Id, 4));
        var result = session.Snapshot.GetSheet(sheet.Id);
        Assert.Contains("A5:C5", result.MergedCells); Assert.Equal("标题", result.GetCell(4, 0)!.Value.TextValue);
        Assert.Null(result.GetCell(4, 9)); Assert.Null(result.GetCell(5, 9));
        session.Undo(); Assert.Equal(sheet.MergedCells, session.Snapshot.GetSheet(sheet.Id).MergedCells);
    }

    [Fact]
    public void CrossSheetCopyCarriesDefaultDimensionsAndInheritedBlankStyles()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Source", defaultRowHeight: 32, columns: [new(4, new(StyleId: 1))]);
        var target = new WorksheetSnapshot(SheetId.New(), "Target", defaultRowHeight: 20);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source, target], styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book);
        var copy = AxisClipboard.Capture(book, source.Id, SheetAxis.Row, 0, 1);
        session.Execute(new InsertCopiedAxisCommand(copy, target.Id, 0));
        Assert.Equal(32, session.Snapshot.GetSheet(target.Id).Rows[0].Height);
        Assert.Equal(1, session.Snapshot.GetSheet(target.Id).GetStyleId(0, 4));
    }

    [Fact]
    public void StaleCopyAndInsertionSplittingSourceLeaveDocumentUntouched()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "S");
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        using var session = new DocumentSession(book);
        var copy = AxisClipboard.Capture(book, sheet.Id, SheetAxis.Row, 1, 2);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new InsertCopiedAxisCommand(copy, sheet.Id, 2)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(1)))]));
        var edited = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new InsertCopiedAxisCommand(copy, sheet.Id, 5)));
        Assert.Same(edited, session.Snapshot);
    }

    [Fact]
    public void ProtectedDestinationAndOverflowRejectEntireInsertion()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Source", [new(new(WorkbookLimits.MaxRows - 1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(1))))]);
        var target = new WorksheetSnapshot(SheetId.New(), "Target", data: new() { Protection = new() { Sheet = true } });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source, target]);
        using var session = new DocumentSession(book);
        var copy = AxisClipboard.Capture(book, source.Id, SheetAxis.Row, 0, 1);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new InsertCopiedAxisCommand(copy, target.Id, 0)));
        Assert.ThrowsAny<Exception>(() => session.Execute(new InsertCopiedAxisCommand(copy, source.Id, 0)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void DifferentAutofitSizesAreAtomicAndPreserveOtherMetadata()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "S", rows: [new(0, new(Height: 25, Hidden: true))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        using var session = new DocumentSession(book);
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 0, 2) { Sizes = new Dictionary<int, double> { [0] = 30, [1] = 50 }.ToImmutableDictionary() });
        Assert.Equal(30, session.Snapshot.Sheets[0].Rows[0].Height); Assert.True(session.Snapshot.Sheets[0].Rows[0].Hidden);
        Assert.Equal(50, session.Snapshot.Sheets[0].Rows[1].Height);
        session.Undo(); Assert.Equal(sheet.Rows, session.Snapshot.Sheets[0].Rows); Assert.False(session.CanUndo);
    }

    [Fact]
    public void RichTextRunsSurviveInsertedRowCopy()
    {
        var rich = new RichCellText([new("red", new(Color: "FFFF0000")), new(" tail")]);
        var source = new WorksheetSnapshot(SheetId.New(), "S", [new(new(0, 0), new(CellContent.FromRichText(rich)))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]);
        using var session = new DocumentSession(book);
        session.Execute(new InsertCopiedAxisCommand(AxisClipboard.Capture(book, source.Id, SheetAxis.Row, 0, 1), source.Id, 2));
        Assert.Equal(rich, session.Snapshot.Sheets[0].GetCell(2, 0)!.Content.RichText);
    }

    [Fact]
    public void ValidationOutsideStoredCellsCannotBeSilentlyDroppedByBoundedCopy()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "S", data: new() { Validations = [new([new(0, 100, 0, 100)], ValidationKind.Whole, formula1: "1", formula2: "2")] });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]);
        Assert.Throws<NotSupportedException>(() => AxisClipboard.Capture(book, source.Id, SheetAxis.Row, 0, 1));
    }

    [Fact]
    public void FailedPasteIntoExpandedMergeLeavesNoInsertedRowsOrHistory()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "S", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("copy"))))], mergedCells: ["A3:A5"]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source]);
        using var session = new DocumentSession(book);
        var copy = AxisClipboard.Capture(book, source.Id, SheetAxis.Row, 0, 1);
        Assert.ThrowsAny<Exception>(() => session.Execute(new InsertCopiedAxisCommand(copy, source.Id, 3)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }
}
