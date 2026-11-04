using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class MergedClipboardTests
{
    private static WorkbookSnapshot MixedBook()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Source",
            [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("merged"), 1))),
             new(new(1, 1), new(CellContent.FromLiteral(CellValue.Blank, 1))),
             new(new(0, 2), T("ordinary")), new(new(2, 2), F("=A1"))],
            mergedCells: ["A1:B2", "C3:D3"]);
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source, new(SheetId.New(), "Target")],
            styles: [new(), new() { Bold = true, FillColor = "FFFFFF00" }]);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CompleteCopyMovesMergeGeometryStylesAndFormulaReferencesWithOneUndo(bool crossSheet, bool transpose)
    {
        var book = MixedBook(); var source = book.Sheets[0]; var target = book.Sheets[crossSheet ? 1 : 0].Id;
        var payload = WorkbookClipboard.Capture(book, source.Id, new(0, 0, 2, 3));
        using var session = new DocumentSession(book);
        var changes = new List<DocumentChange>(); session.Changed += (_, change) => changes.Add(change);
        session.Execute(new PasteCellsCommand(payload, target, new(4, 6, 4, 6), 0, transpose: transpose));
        var result = session.Snapshot.GetSheet(target);
        Assert.Contains("G5:H6", result.MergedCells);
        Assert.Contains(transpose ? "I7:I8" : "I7:J7", result.MergedCells);
        Assert.Equal("merged", result.GetCell(4, 6)!.Value.TextValue);
        Assert.Equal(1, result.GetStyleId(5, 7));
        Assert.Equal("ordinary", result.GetCell(transpose ? 6 : 4, transpose ? 6 : 8)!.Value.TextValue);
        Assert.Equal("=G5", result.GetCell(6, 8)!.Content.FormulaText);
        Assert.True(Assert.Single(changes).StructureChanged);
        RoundTrip(session.Snapshot, target);
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal(book.GetSheet(target).MergedCells.ToArray(), session.Snapshot.GetSheet(target).MergedCells.ToArray());
        Assert.Null(session.Snapshot.GetSheet(target).GetCell(4, 6));
        Assert.True(session.Redo()); Assert.Equal(result.MergedCells.ToArray(), session.Snapshot.GetSheet(target).MergedCells.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormatPasteMergesAndUnmergesFilledTargetsWithoutLosingCoveredValuesOrFormulas(bool scalarUnmerge)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data",
            [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("source"), 1))),
             new(new(3, 3), T("target")), new(new(3, 4), F("=2+3", 5)),
             new(new(4, 3), T("below")), new(new(4, 4), N(123))],
            mergedCells: ["A1:B2"],
            annotations: new(links: [new(new(3, 3), new("https://example.com"))]));
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet],
            styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, sheet.Id, new(0, 0, 1, 1));
        var merged = session.Execute(new PasteCellsCommand(clipboard, sheet.Id, new(3, 3, 4, 4), 0, CellPasteMode.Formats));
        Assert.True(merged.Changed); Assert.False(merged.CalculationUnchanged);
        Assert.Contains("D4:E5", session.Snapshot.Sheets[0].MergedCells);
        AssertTargetContents(); Assert.Equal(1, session.Snapshot.Sheets[0].GetStyleId(3, 3));
        RoundTrip(session.Snapshot, sheet.Id);
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.DoesNotContain("D4:E5", session.Snapshot.Sheets[0].MergedCells); AssertTargetContents();
        Assert.True(session.Redo());
        var plain = WorkbookClipboard.Capture(session.Snapshot, sheet.Id, scalarUnmerge ? new(0, 10, 0, 10) : new(0, 10, 1, 11));
        session.Execute(new PasteCellsCommand(plain, sheet.Id, new(3, 3, 4, 4), session.Revision, CellPasteMode.Formats));
        Assert.DoesNotContain("D4:E5", session.Snapshot.Sheets[0].MergedCells);
        AssertTargetContents(); Assert.Equal(0, session.Snapshot.Sheets[0].GetStyleId(3, 3));
        RoundTrip(session.Snapshot, sheet.Id);
        Assert.True(session.Undo()); Assert.Contains("D4:E5", session.Snapshot.Sheets[0].MergedCells); AssertTargetContents();
        Assert.True(session.Redo()); Assert.DoesNotContain("D4:E5", session.Snapshot.Sheets[0].MergedCells); AssertTargetContents();

        void AssertTargetContents()
        {
            var current = session.Snapshot.Sheets[0];
            foreach (var p in sheet.Cells.Where(p => p.Key.Row >= 3))
                Assert.Equal(p.Value.Content.WithStyle(0), current.GetCell(p.Key.Row, p.Key.Column)!.Content.WithStyle(0));
            Assert.Same(sheet.Annotations, current.Annotations);
        }
    }

    [Theory]
    [InlineData(CellPasteMode.All)]
    [InlineData(CellPasteMode.Formats)]
    public void PasteReplacesExistingMergeLayoutWithinTheWholeTarget(CellPasteMode mode)
    {
        var book = MixedBook(); var target = new WorksheetSnapshot(book.Sheets[1].Id, "Target", mergedCells: ["G5:J5", "L1:M1"]);
        book = book.WithSheets([book.Sheets[0], target]);
        using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, book.Sheets[0].Id, new(0, 0, 2, 3));
        session.Execute(new PasteCellsCommand(clipboard, target.Id, new(4, 6, 6, 9), 0, mode));
        Assert.Equal(new[] { "L1:M1", "G5:H6", "I7:J7" }, session.Snapshot.GetSheet(target.Id).MergedCells.ToArray());
        Assert.True(session.Undo()); Assert.Equal(target.MergedCells.ToArray(), session.Snapshot.GetSheet(target.Id).MergedCells.ToArray());
    }

    [Theory]
    [InlineData(CellPasteMode.Values, false)]
    [InlineData(CellPasteMode.Values, true)]
    [InlineData(CellPasteMode.Formulas, false)]
    [InlineData(CellPasteMode.Formulas, true)]
    public void ContentOnlyPastePreservesDestinationMergeLayout(CellPasteMode mode, bool mergedTarget)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data", [new(new(0, 0), N(7))],
            mergedCells: mergedTarget ? ["A1:B2", "D4:E5"] : ["A1:B2"]);
        var book = Book(source); using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, source.Id, new(0, 0, 1, 1));
        session.Execute(new PasteCellsCommand(clipboard, source.Id, new(3, 3, 4, 4), 0, mode));
        Assert.Equal(source.MergedCells.ToArray(), session.Snapshot.Sheets[0].MergedCells.ToArray());
        Assert.Equal(CellValue.FromNumber(7), session.Snapshot.Sheets[0].GetCell(3, 3)!.Value);
    }

    [Theory]
    [InlineData(CellPasteMode.All)]
    [InlineData(CellPasteMode.Formats)]
    public void CopyOfAlreadyMergedCoveredInputsPreservesThemAndNoOpDoesNotAddHistory(CellPasteMode mode)
    {
        var source = new WorksheetSnapshot(SheetId.New(), "Data",
            [new(new(0, 0), T("anchor")), new(new(0, 1), F("=2+3"))], mergedCells: ["A1:B2", "D4:E5"]);
        var book = Book(source); using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, source.Id, new(0, 0, 1, 1));
        Assert.False(session.Execute(new PasteCellsCommand(clipboard, source.Id, new(0, 0, 1, 1), 0, mode)).Changed);
        Assert.False(session.CanUndo);
        session.Execute(new PasteCellsCommand(clipboard, source.Id, new(3, 3, 4, 4), 0));
        Assert.Equal("=2+3", session.Snapshot.Sheets[0].GetCell(3, 4)!.Content.FormulaText);
        Assert.Equal(source.MergedCells.ToArray(), session.Snapshot.Sheets[0].MergedCells.ToArray());
        RoundTrip(session.Snapshot, source.Id);
    }

    [Fact]
    public void PartialMergeCaptureAndCutKeepTheirExplicitGuards()
    {
        var book = MixedBook(); var id = book.Sheets[0].Id;
        Assert.Throws<InvalidOperationException>(() => WorkbookClipboard.Capture(book, id, new(0, 0, 0, 1)));
        Assert.Throws<InvalidOperationException>(() => WorkbookClipboard.Capture(book, id, new(1, 1, 2, 3)));
        Assert.Throws<NotSupportedException>(() => WorkbookClipboard.Capture(book, id, new(0, 0, 1, 1), cut: true));
    }

    [Theory]
    [InlineData(CellPasteMode.All)]
    [InlineData(CellPasteMode.Formats)]
    public void PartialTargetMergeRejectsBeforeChangingContentsOrGeometry(CellPasteMode mode)
    {
        var book = MixedBook(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, id, new(2, 2, 2, 3));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(clipboard, id, new(0, 1, 0, 1), 0, mode)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MergeFormattingDoesNotHideAttachmentsOrBypassSheetProtection(bool protection)
    {
        var book = MixedBook(); var source = book.Sheets[0];
        var target = new WorksheetSnapshot(book.Sheets[1].Id, "Target",
            annotations: protection ? null : new(notes: [new(new(4, 7), new("Author", new([new("Keep")])))]),
            data: protection ? new() { Protection = new() } : null);
        book = book.WithSheets([source, target]); using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, source.Id, new(0, 0, 1, 1));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(clipboard, target.Id, new(4, 6, 4, 6), 0, CellPasteMode.Formats)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void MergeFormatChangesRecalculateSpillObstaclesAndUndoRestoresTheArray()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", [new(new(4, 4), F("=SEQUENCE(2,2)"))], mergedCells: ["A1:B2"]);
        using var doc = new WorkbookDocument(Book(sheet), new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        doc.Calculate(); Assert.Equal("4", doc.GetCellValue(0, 5, 5));
        var clipboard = WorkbookClipboard.Capture(doc.Snapshot, sheet.Id, new(0, 0, 1, 1));
        doc.Session.Execute(new PasteCellsCommand(clipboard, sheet.Id, new(4, 4, 5, 5), doc.Snapshot.Revision, CellPasteMode.Formats));
        doc.Calculate();
        Assert.Equal(CellValue.FromError(CellErrorCode.Spill), doc.Snapshot.Sheets[0].GetCell(4, 4)!.Value);
        Assert.True(doc.Undo()); doc.Calculate();
        Assert.Equal("4", doc.GetCellValue(0, 5, 5)); Assert.DoesNotContain("E5:F6", doc.Snapshot.Sheets[0].MergedCells);
    }

    private static void RoundTrip(WorkbookSnapshot book, SheetId target)
    {
        var directory = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "merged-clipboard", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory); var storage = new TypedWorkbookStorage(); var expected = book.GetSheet(target);
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "roundtrip." + extension); storage.Save(path, book);
            var reopened = storage.Open(path).Sheets.Single(s => s.Name == expected.Name);
            Assert.Equal(expected.MergedCells.ToArray(), reopened.MergedCells.ToArray());
            Assert.Equal(expected.Cells.ToDictionary(p => p.Key, p => p.Value.Content), reopened.Cells.ToDictionary(p => p.Key, p => p.Value.Content));
            Assert.True(expected.Annotations.ContentEquals(reopened.Annotations));
        }
    }
}
