using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class PasteOptionsTests
{
    private static WorkbookSnapshot Book() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new(SheetId.New(), "Sheet1", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(12), 1))),
            new(new(0, 1), new(CellContent.FromLiteral(CellValue.FromNumber(34), 1))),
            new(new(3, 3), new(CellContent.FromLiteral(CellValue.FromText("original"), 2)))])],
        styles: [new(), new() { Bold = true }, new() { Italic = true }]);

    [Theory]
    [InlineData(CellPasteMode.Values)]
    [InlineData(CellPasteMode.Formats)]
    [InlineData(CellPasteMode.Formulas)]
    public void ChangingModeUsesOriginalDestinationAndOneUndo(CellPasteMode mode)
    {
        var book = Book(); var sheet = book.ActiveSheetId; using var session = new DocumentSession(book);
        var source = WorkbookClipboard.Capture(book, sheet, new(0, 0, 0, 1));
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), 0), 0);
        var changes = new List<DocumentChange>(); session.Changed += (_, change) => changes.Add(change);
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), 1, mode), 1, replacePrevious: true);
        var cell = session.Snapshot.Sheets[0].GetCell(3, 3)!;
        Assert.Equal(mode == CellPasteMode.Formats ? CellValue.FromText("original") : CellValue.FromNumber(12), cell.Value);
        Assert.Equal(mode == CellPasteMode.Formats ? 1 : 2, cell.Content.StyleId);
        Assert.Single(changes); Assert.Equal(2, session.Revision);
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal(book.Sheets[0].GetCell(3, 3)!.Content, session.Snapshot.Sheets[0].GetCell(3, 3)!.Content);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(3, 4));
        Assert.True(session.Redo()); Assert.Equal(cell.Content, session.Snapshot.Sheets[0].GetCell(3, 3)!.Content);
    }

    [Fact]
    public void TransposeRestoresCellsOutsideTheNewFootprint()
    {
        var book = Book(); var sheet = book.ActiveSheetId; using var session = new DocumentSession(book);
        var source = WorkbookClipboard.Capture(book, sheet, new(0, 0, 0, 1));
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), 0), 0);
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), 1, transpose: true), 1, true);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(3, 4));
        Assert.Equal(CellValue.FromNumber(34), session.Snapshot.Sheets[0].GetCell(4, 3)!.Value);
        session.Undo(); Assert.Null(session.Snapshot.Sheets[0].GetCell(4, 3)); Assert.False(session.CanUndo);
    }

    [Fact]
    public void FailedReplacementLeavesPasteAndHistoryIntact()
    {
        var book = Book(); var sheet = book.ActiveSheetId; using var session = new DocumentSession(book);
        var source = WorkbookClipboard.Capture(book, sheet, new(0, 0, 0, 1));
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), 0), 0);
        var pasted = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 5, 5), 1), 1, true));
        Assert.Same(pasted, session.Snapshot); Assert.False(session.CanRedo);
        session.Undo(); Assert.False(session.CanUndo);
    }

    [Theory]
    [InlineData("edit")]
    [InlineData("save")]
    [InlineData("undo")]
    public void ChangedSavedOrUndonePasteCannotBeReplaced(string action)
    {
        var book = Book(); var sheet = book.ActiveSheetId; using var session = new DocumentSession(book);
        var source = WorkbookClipboard.Capture(book, sheet, new(0, 0, 0, 1));
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), 0), 0);
        if (action == "edit") session.Execute(new SetCellsCommand([new(new(sheet, 8, 8), CellContent.FromLiteral(CellValue.FromNumber(99)))]));
        if (action == "save") session.MarkSaved(session.SessionId, session.Revision);
        if (action == "undo") session.Undo();
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.ExecuteCellPaste(new PasteCellsCommand(source, sheet, new(3, 3, 3, 3), session.Revision), session.Revision, true));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void ExternalValuesRemoveNewMergeAndRestoreOriginalStyles()
    {
        var book = Book(); var sheet = book.ActiveSheetId; using var session = new DocumentSession(book);
        var source = new ExternalClipboardRegion(1, 2, [new(CellValue.FromNumber(5), new() { Bold = true }), new(CellValue.Blank, new())],
            merges: [new(0, 0, 0, 1)]);
        session.ExecuteCellPaste(new PasteExternalCellsCommand(source, sheet, new(3, 3, 3, 3), 0), 0);
        Assert.Single(session.Snapshot.Sheets[0].MergedCells);
        session.ExecuteCellPaste(new PasteExternalCellsCommand(source, sheet, new(3, 3, 3, 3), 1, CellPasteMode.Values), 1, true);
        Assert.Empty(session.Snapshot.Sheets[0].MergedCells);
        Assert.Equal(2, session.Snapshot.Sheets[0].GetCell(3, 3)!.Content.StyleId);
        session.Undo(); Assert.Equal(CellValue.FromText("original"), session.Snapshot.Sheets[0].GetCell(3, 3)!.Value);
    }

    [Fact]
    public void ReplacementStillValidatesAndDoesNotPublishOnFailure()
    {
        var book = Book(); var sheet = book.Sheets[0];
        var rule = new WorksheetValidation([new(4, 3, 4, 3)], ValidationKind.Whole, "=1", "=20");
        book = book.WithSheets([sheet.WithData(new() { Validations = [rule] })]);
        using var session = new DocumentSession(book, validation: new Calculation.WorkbookValidation());
        var source = WorkbookClipboard.Capture(book, sheet.Id, new(0, 0, 0, 1));
        session.ExecuteCellPaste(new PasteCellsCommand(source, sheet.Id, new(3, 3, 3, 3), 0), 0);
        var before = session.Snapshot; var notifications = 0; session.Changed += (_, _) => notifications++;
        Assert.Throws<DataValidationException>(() => session.ExecuteCellPaste(
            new PasteCellsCommand(source, sheet.Id, new(3, 3, 3, 3), 1, transpose: true), 1, true));
        Assert.Same(before, session.Snapshot); Assert.Equal(0, notifications);
        session.Undo(); Assert.False(session.CanUndo);
    }
}
