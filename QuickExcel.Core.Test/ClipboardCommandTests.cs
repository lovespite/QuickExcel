using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ClipboardCommandTests
{
    private static WorkbookSnapshot Book() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new(SheetId.New(), "Source", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("00123"), 1))),
            new(new(0, 1), new(CellContent.FromFormula("=A3+$A$3+$A3+A$3", 1), new(CellValue.FromNumber(40), CalculationState.Fresh, 0, 0, 0))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            new(new(4, 4), new(CellContent.FromLiteral(CellValue.FromText("old"), 2)))]), new(SheetId.New(), "Target")],
        styles: [new(), new() { Bold = true, NumberFormatId = 49 }, new() { Italic = true, NumberFormatId = 2 }]);

    [Theory]
    [InlineData(CellPasteMode.All)]
    [InlineData(CellPasteMode.Values)]
    [InlineData(CellPasteMode.Formulas)]
    [InlineData(CellPasteMode.Formats)]
    public void PasteModesHaveOneUndoPreserveTypesAndRoundTrip(CellPasteMode mode)
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 1, 1));
        session.Execute(new PasteCellsCommand(payload, id, new(4, 4, 4, 4), 0, mode));
        var result = session.Snapshot; var sheet = result.Sheets[0]; var first = sheet.GetCell(4, 4)!;
        Assert.Equal(mode == CellPasteMode.Formats ? "old" : "00123", first.Value.TextValue);
        Assert.Equal(mode is CellPasteMode.All or CellPasteMode.Formats ? 1 : 2, first.Content.StyleId);
        if (mode is CellPasteMode.All or CellPasteMode.Formulas)
            Assert.Equal("=(((E7+$A$3)+$A7)+E$3)", sheet.GetCell(4, 5)!.Content.FormulaText);
        else if (mode == CellPasteMode.Values) Assert.Equal(CellValue.FromNumber(40), sheet.GetCell(4, 5)!.Value);
        if (mode != CellPasteMode.Formats) Assert.Equal(CellValue.FromBoolean(true), sheet.GetCell(5, 4)!.Value);
        Assert.True(session.Undo()); EqualContent(book.Sheets[0], session.Snapshot.Sheets[0]); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); EqualContent(sheet, session.Snapshot.Sheets[0]);
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "clipboard-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(folder, "paste." + extension); var storage = new TypedWorkbookStorage(); storage.Save(path, result);
            var reopened = storage.Open(path); Assert.Equal(first.Content, reopened.Sheets[0].GetCell(4, 4)!.Content);
            Assert.Equal(sheet.GetCell(4, 5)?.Content, reopened.Sheets[0].GetCell(4, 5)?.Content);
        }
    }

    [Theory]
    [InlineData("=A6", 0, 0, 6, 4, "=J7")]
    [InlineData("=A$6", 0, 1, 7, 4, "=A$6")]
    [InlineData("=$A6", 1, 0, 6, 5, "=$A6")]
    [InlineData("=$A$6", 1, 1, 7, 5, "=$A$6")]
    [InlineData("=SUM(A6:B8)", 4, 0, 6, 8, "=SUM(J7:L8)")]
    [InlineData("=SUM($A6:B$8)", 4, 1, 7, 8, "=SUM($A6:B$8)")]
    public void TransposeMatchesIndependentlyObservedExcelReferences(string formula, int sr, int sc, int dr, int dc, string expected) =>
        Assert.Equal(expected, FormulaCopy.Translate(formula, new(sr, sc), new(dr, dc), true));

    [Fact]
    public void CopySupportsWholeColumnsQualifiedReferencesAndRefErrors()
    {
        Assert.Equal("=SUM('Other'!C:C,$A:$A)", FormulaCopy.Translate("=SUM(Other!A:A,$A:$A)", new(1, 1), new(4, 3)));
        Assert.Equal("=#REF!", FormulaCopy.Translate("=A1", new(1, 1), new(0, 0)));
        Assert.Throws<NotSupportedException>(() => FormulaCopy.Translate("=SUM(A1:B2)", new(1, 1), new(0, 0)));
        Assert.Throws<NotSupportedException>(() => FormulaCopy.Translate("=SUM(A:A)", new(0, 0), new(2, 2), true));
    }

    [Theory]
    [InlineData("A3:B4", "G5:H6")]
    [InlineData("$A3:B4", "G$5:H6")]
    [InlineData("A$3:B4", "$G5:H6")]
    [InlineData("$A$3:B4", "$G$5:H6")]
    [InlineData("A3:$B4", "G5:H$6")]
    [InlineData("$A3:$B4", "$A3:$B4")]
    [InlineData("A$3:$B4", "A$3:$B4")]
    [InlineData("$A$3:$B4", "$A$3:$B4")]
    [InlineData("A3:B$4", "G5:$H6")]
    [InlineData("$A3:B$4", "$A3:B$4")]
    [InlineData("A$3:B$4", "A$3:B$4")]
    [InlineData("$A$3:B$4", "$A$3:B$4")]
    [InlineData("A3:$B$4", "G5:$H$6")]
    [InlineData("$A3:$B$4", "$A3:$B$4")]
    [InlineData("A$3:$B$4", "A$3:$B$4")]
    [InlineData("$A$3:$B$4", "$A$3:$B$4")]
    public void TransposeRangeDollarFlagsMatchAllSixteenExcelCombinations(string source, string target) =>
        Assert.Equal("=SUM(" + target + ")", FormulaCopy.Translate("=SUM(" + source + ")", new(0, 2), new(6, 4), true));

    [Fact]
    public void TransposeAndScalarFillMapOnlySelectedCells()
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 0, 1));
        session.Execute(new PasteCellsCommand(payload, id, new(4, 4, 4, 4), 0, transpose: true));
        Assert.Equal("00123", session.Snapshot.Sheets[0].GetCell(4, 4)!.Value.TextValue);
        Assert.Equal("=(((G5+$A$3)+$A3)+A$3)", session.Snapshot.Sheets[0].GetCell(5, 4)!.Content.FormulaText);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(4, 5));
        var scalar = WorkbookClipboard.Capture(session.Snapshot, id, new(0, 1, 0, 1));
        session.Execute(new PasteCellsCommand(scalar, id, new(8, 0, 9, 1), 1));
        Assert.Contains("$A11", session.Snapshot.Sheets[0].GetCell(8, 1)!.Content.FormulaText);
        Assert.Contains("$A12", session.Snapshot.Sheets[0].GetCell(9, 1)!.Content.FormulaText);
    }

    [Fact]
    public void RepeatedAndOverlappingCopyUsesCapturedInputsAndExpiresAfterAnotherEdit()
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 0, 1));
        session.Execute(new PasteCellsCommand(payload, id, new(0, 1, 0, 1), 0));
        Assert.Equal("00123", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        session.Execute(new PasteCellsCommand(payload, id, new(1, 1, 1, 1), 1));
        Assert.Equal(CellContentKind.Formula, session.Snapshot.Sheets[0].GetCell(1, 2)!.Content.Kind);
        var afterPaste = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(payload, id, new(2, 1, 2, 1), 0)));
        Assert.Same(afterPaste, session.Snapshot);
        session.Execute(new SetCellsCommand([new(new(id, 9, 9), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(payload, id, new(2, 1, 2, 1), 2)));
    }

    [Fact]
    public void OverlappingCutMovesReferencesAndAllCellsInOneUndo()
    {
        var id = SheetId.New(); var sheet = new WorksheetSnapshot(id, "S", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("00123"), 1))),
            new(new(0, 1), new(CellContent.FromLiteral(CellValue.FromNumber(2)))), new(new(0, 3), new(CellContent.FromFormula("=$A$1+B1")))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book); var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 0, 1), cut: true);
        Assert.False(session.IsDirty); Assert.False(session.CanUndo);
        session.Execute(new PasteCellsCommand(payload, id, new(0, 1, 0, 1), 0));
        Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 0)); Assert.Equal("00123", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        Assert.Equal(2, session.Snapshot.Sheets[0].GetCell(0, 2)!.Value.NumberValue);
        Assert.Equal("=('S'!$B$1+'S'!C1)", session.Snapshot.Sheets[0].GetCell(0, 3)!.Content.FormulaText);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(payload, id, new(2, 0, 2, 0), 1)));
        session.Undo(); EqualContent(sheet, session.Snapshot.Sheets[0]); Assert.False(session.CanUndo);
    }

    [Theory]
    [InlineData(CellPasteMode.Values, false)]
    [InlineData(CellPasteMode.Formulas, false)]
    [InlineData(CellPasteMode.Formats, false)]
    [InlineData(CellPasteMode.All, true)]
    public void CutRejectsSelectivePasteAtomically(CellPasteMode mode, bool transpose)
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 1, 1), true);
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(payload, id, new(4, 4, 4, 4), 0, mode, transpose)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void RichRunsAndAttachmentsAreNeverSilentlyLost()
    {
        var id = SheetId.New(); var rich = new RichCellText([new("red", new(Color: "FFFF0000")), new(" tail")]);
        var sheet = new WorksheetSnapshot(id, "S", [new(new(0, 0), new(CellContent.FromRichText(rich))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromText("plain"))))],
            annotations: new([new(new(2, 0), new("https://example.invalid/"))], [new(new(3, 0), new("author", new([new("note")])))]));
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]); using var session = new DocumentSession(book);
        var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 0, 0));
        session.Execute(new PasteCellsCommand(payload, id, new(0, 2, 0, 2), 0));
        Assert.Equal(rich, session.Snapshot.Sheets[0].GetCell(0, 2)!.Content.RichText);
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(payload, id, new(0, 3, 0, 3), 1, CellPasteMode.Formats)));
        var plain = WorkbookClipboard.Capture(session.Snapshot, id, new(1, 0, 1, 0));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(plain, id, new(0, 2, 0, 2), 1, CellPasteMode.Values)));
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(plain, id, new(2, 0, 2, 0), 1)));
        session.Execute(new PasteCellsCommand(plain, id, new(2, 0, 2, 0), 1, CellPasteMode.Values));
        Assert.True(sheet.Annotations.ContentEquals(session.Snapshot.Sheets[0].Annotations));
    }

    [Fact]
    public void HiddenMergedOversizeStaleAndMissingResultsRejectWithoutPartialWrites()
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 1, 1));
        Assert.Throws<ArgumentException>(() => session.Execute(new PasteCellsCommand(payload, id, new(2, 2, 4, 4), 0)));
        Assert.Throws<ArgumentException>(() => WorkbookClipboard.Capture(book, id, new(0, 0, 20000, 1)));
        session.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 7, 1, hidden: true));
        var current = session.Snapshot; var newer = WorkbookClipboard.Capture(current, id, new(0, 0, 1, 1));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(newer, id, new(6, 0, 6, 0), 1)));
        Assert.Same(current, session.Snapshot);
        session.Execute(new MergeCellsCommand(id, new(8, 0, 8, 1)));
        Assert.Throws<NotSupportedException>(() => WorkbookClipboard.Capture(session.Snapshot, id, new(8, 0, 8, 1)));
        session.Execute(new SetCellsCommand([new(new(id, 9, 0), CellContent.FromFormula("=1+2"))]));
        var missing = WorkbookClipboard.Capture(session.Snapshot, id, new(9, 0, 9, 0));
        Assert.True(missing.HasOldResults); current = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(missing, id, new(10, 0, 10, 0), current.Revision, CellPasteMode.Values)));
        Assert.Same(current, session.Snapshot);
    }

    [Fact]
    public void CopyBlankEffectiveStyleOverridesDestinationInheritance()
    {
        var id = SheetId.New(); var sheet = new WorksheetSnapshot(id, "S", columns: [new(2, new(StyleId: 1))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }]);
        using var session = new DocumentSession(book);
        session.Execute(new PasteCellsCommand(WorkbookClipboard.Capture(book, id, new(0, 0, 0, 0)), id, new(0, 2, 0, 2), 0));
        Assert.Equal(0, session.Snapshot.Sheets[0].GetStyleId(0, 2)); Assert.NotNull(session.Snapshot.Sheets[0].GetCell(0, 2));
        session.Undo(); Assert.Equal(1, session.Snapshot.Sheets[0].GetStyleId(0, 2));
    }

    [Fact]
    public void LocalNamesAndForeignWorkbookCannotAcquireDifferentMeaningOnPaste()
    {
        var book = Book(); var id = book.Sheets[0].Id;
        book = new(book.DocumentId, 0, book.DateSystem, book.Sheets, styles: book.Styles,
            definedNames: [new("LocalTotal", "=Source!$A$3", id)]);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(id, 2, 2), CellContent.FromFormula("=LocalTotal"))]));
        var payload = WorkbookClipboard.Capture(session.Snapshot, id, new(2, 2, 2, 2)); var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(payload, book.Sheets[1].Id, new(0, 0, 0, 0), 1)));
        Assert.Same(before, session.Snapshot);
        using var other = new DocumentSession(Book());
        Assert.Throws<InvalidOperationException>(() => other.Execute(new PasteCellsCommand(payload, other.Snapshot.Sheets[0].Id, new(0, 0, 0, 0), 0)));
    }

    [Fact]
    public void CutDoesNotDropDestinationRichTextOrPartiallyMoveReferencedRanges()
    {
        var book = Book(); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(id, 9, 0), CellContent.FromRichText(new([new("rich", new(Bold: true))])))]));
        var payload = WorkbookClipboard.Capture(session.Snapshot, id, new(0, 0, 0, 0), true); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(payload, id, new(9, 0, 9, 0), 1)));
        Assert.Same(before, session.Snapshot);
        session.Execute(new SetCellsCommand([new(new(id, 5, 0), CellContent.FromFormula("=SUM(A1:B2)"))]));
        payload = WorkbookClipboard.Capture(session.Snapshot, id, new(0, 0, 0, 0), true); before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(payload, id, new(9, 1, 9, 1), 2)));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void ImportedUnknownFunctionAllowsExplicitCachePasteButCannotBeRelocated()
    {
        var book = Book(); var id = book.Sheets[0].Id;
        var sheets = book.Sheets.SetItem(0, new(id, "Source", [new(new(0, 0), new(CellContent.FromFormula("=UNKNOWN(A1)"),
            new(CellValue.FromNumber(99), CalculationState.Stale, 0, 0, 0)))]));
        book = new(book.DocumentId, 0, book.DateSystem, sheets, styles: book.Styles);
        book = new(book.DocumentId, 0, book.DateSystem, sheets, styles: book.Styles,
            fileState: new("unknown.xlsx", 2, [], book, [], []));
        using var session = new DocumentSession(book); var payload = WorkbookClipboard.Capture(book, id, new(0, 0, 0, 0));
        Assert.True(payload.HasOldResults);
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(payload, id, new(4, 4, 4, 4), 0)));
        Assert.Same(book, session.Snapshot);
        session.Execute(new PasteCellsCommand(payload, id, new(4, 4, 4, 4), 0, CellPasteMode.Values));
        Assert.Equal(CellValue.FromNumber(99), session.Snapshot.Sheets[0].GetCell(4, 4)!.Value);
        Assert.Equal(CalculationState.Stale, session.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
    }
    private static void EqualContent(WorksheetSnapshot expected, WorksheetSnapshot actual) =>
        Assert.Equal(expected.Cells.ToDictionary(p => p.Key, p => p.Value.Content), actual.Cells.ToDictionary(p => p.Key, p => p.Value.Content));
}
