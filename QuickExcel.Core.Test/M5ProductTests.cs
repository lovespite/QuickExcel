using QuickExcel.Domain;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M5ProductTests
{
    [Fact]
    public void SearchUsesSparseRangePredicateAndDoesNotRevealProtectedFormulas()
    {
        var sheet = Sheet("Data", (0, 0, F("=1+1")), (WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1, T("far")));
        using var session = new DocumentSession(Book(sheet));
        var calls = 0;
        var hits = WorkbookSearch.Find(session.Snapshot, 0, new("far", SearchScope.Selection), selected: address =>
        { calls++; return address.SheetId == sheet.Id; });
        Assert.Single(hits); Assert.Equal(2, calls);
        session.Execute(new SetCellProtectionCommand(sheet.Id, new(0, 0, 0, 0), true, true));
        session.Execute(new SetProtectionCommand(sheet.Id, new()));
        Assert.Empty(WorkbookSearch.Find(session.Snapshot, 0, new("1", Target: SearchTarget.Formula)));
        session.Execute(new SetProtectionCommand(sheet.Id, null));
        Assert.Single(WorkbookSearch.Find(session.Snapshot, 0, new("1", Target: SearchTarget.Formula)));
    }

    [Fact]
    public void WorkbookSearchAndReplacementExcludeHiddenWorksheets()
    {
        var visible = Sheet("Visible", (0, 0, T("hit")));
        var hidden = Sheet("Hidden", (0, 0, T("hit"))).WithVisibility(SheetVisibility.Hidden);
        var veryHidden = Sheet("VeryHidden", (0, 0, T("hit"))).WithVisibility(SheetVisibility.VeryHidden);
        using var session = new DocumentSession(Book(visible, hidden, veryHidden));
        var options = new SearchOptions("hit", SearchScope.Workbook, IncludeHidden: true);
        var hits = WorkbookSearch.Find(session.Snapshot, 0, options);
        Assert.Single(hits);
        session.Execute(WorkbookSearch.Replace(hits, options, "changed"));
        Assert.Equal(CellValue.FromText("changed"), session.Snapshot.GetSheet(visible.Id).GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText("hit"), session.Snapshot.GetSheet(hidden.Id).GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText("hit"), session.Snapshot.GetSheet(veryHidden.Id).GetCell(0, 0)!.Value);
    }

    [Fact]
    public void ReplaceAcrossSheetsIsOneTransactionAndPreservesTextIdentity()
    {
        var first = Sheet("First", (0, 0, T("apple")), (0, 1, T("apple")));
        var second = Sheet("Second", (0, 0, T("APPLE")));
        using var session = new DocumentSession(Book(first, second));
        var options = new SearchOptions("apple", SearchScope.Workbook);
        var matches = WorkbookSearch.Find(session.Snapshot, 0, options);
        Assert.Equal(3, matches.Count);
        session.Execute(WorkbookSearch.Replace(matches, options, "=1+1"));
        Assert.Equal(1, session.Revision);
        Assert.All(session.Snapshot.Sheets.SelectMany(s => s.Cells.Values), c => Assert.Equal(CellValue.FromText("=1+1"), c.Value));
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal("APPLE", session.Snapshot.Sheets[1].GetCell(0, 0)!.Value.TextValue);
    }

    [Fact]
    public void SearchSupportsDiscontiguousSelectionHiddenAndFilteredPolicy()
    {
        var sheet = Sheet("Data", (0, 0, T("hit")), (1, 0, T("hit")), (2, 0, T("hit")), (3, 0, T("hit")), (0, 1, T("hit")));
        sheet = new(sheet.Id, sheet.Name, sheet.Cells, rows: new Dictionary<int, RowMetadata> { [1] = new(Hidden: true) }, columns: new Dictionary<int, ColumnMetadata> { [1] = new(Hidden: true) });
        var book = Book(sheet);
        Assert.Equal(2, WorkbookSearch.Find(book, 0, new("hit"), filtered: (_, r) => r == 2).Count);
        Assert.Equal(5, WorkbookSearch.Find(book, 0, new("hit", IncludeHidden: true)).Count);
        var selected = new HashSet<CellAddress> { new(sheet.Id, 0, 0), new(sheet.Id, 3, 0) };
        Assert.Equal(2, WorkbookSearch.Find(book, 0, new("hit", SearchScope.Selection), selected).Count);
        Assert.Single(WorkbookSearch.Find(book, 0, new("hit", MatchCase: true), new HashSet<CellAddress>()), h => h.Address.Row == 0);
    }

    [Fact]
    public void ReplaceDisplayFormulaResultRejectsEntireOperation()
    {
        var formula = new CellSnapshot(CellContent.FromFormula("=1"), new(CellValue.FromNumber(1), CalculationState.Fresh, 0, 0, 0));
        using var session = new DocumentSession(Book(Sheet("Data", (0, 0, N(1)), (1, 0, formula))));
        var before = session.Snapshot;
        var options = new SearchOptions("1");
        Assert.Throws<InvalidOperationException>(() => session.Execute(WorkbookSearch.Replace(WorkbookSearch.Find(before, 0, options), options, "2")));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void FormulaSearchAndReplaceValidatesAllBeforeWriting()
    {
        using var session = new DocumentSession(Book(Sheet("Data", (0, 0, F("=A2+1")), (0, 1, F("=A2*2")))));
        var options = new SearchOptions("A2", Target: SearchTarget.Formula);
        var matches = WorkbookSearch.Find(session.Snapshot, 0, options);
        Assert.Throws<FormatException>(() => WorkbookSearch.Replace(matches, options, "("));
        Assert.False(session.CanUndo);
        session.Execute(WorkbookSearch.Replace(matches, options, "B2"));
        Assert.Equal("=B2+1", session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
        session.Undo(); Assert.Equal("=A2*2", session.Snapshot.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
    }

    [Fact]
    public void SearchCaseWholeCellAndCapacityAreExplicit()
    {
        var book = Book(Sheet("Data", (0, 0, T("Apple")), (1, 0, T("apple pie"))));
        Assert.Single(WorkbookSearch.Find(book, 0, new("Apple", MatchCase: true)));
        Assert.Single(WorkbookSearch.Find(book, 0, new("apple", WholeCell: true)));
        Assert.Throws<ArgumentException>(() => WorkbookSearch.Find(book, 0, new("")));
        var many = new WorksheetSnapshot(SheetId.New(), "Many", Enumerable.Range(0, WorkbookSearch.Limit + 1).Select(i => new KeyValuePair<CellPosition, CellSnapshot>(new(i, 0), T("x"))));
        Assert.Throws<InvalidOperationException>(() => WorkbookSearch.Find(Book(many), 0, new("x")));
        Assert.Throws<ArgumentException>(() => WorkbookSearch.Replace(WorkbookSearch.Find(book, 0, new("Apple")), new("Apple"), new string('x', 32768)));
    }

    [Fact]
    public void ReplacementExpansionIsBoundedBeforeAllocationAndPreservesTheDocument()
    {
        using var session = new DocumentSession(Book(Sheet("Data", (0, 0, T(new string('x', WorkbookLimits.MaxCellTextLength))))));
        var before = session.Snapshot;
        var options = new SearchOptions("x");
        var hits = WorkbookSearch.Find(before, 0, options);
        Assert.Throws<ArgumentException>(() => session.Execute(WorkbookSearch.Replace(hits, options, new string('y', WorkbookLimits.MaxCellTextLength))));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void StructureTranslatesMergeImageAnchorsAndRestoresOnUndo(SheetAxis axis)
    {
        var image = new SheetImage([1, 2, 3], "png", 2, 2, 2, 2);
        var sheet = new WorksheetSnapshot(SheetId.New(), "Art", mergedCells: ["C3:D4"], images: [image]);
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new ShiftAxisCommand(sheet.Id, axis, 1, 2));
        var current = session.Snapshot.Sheets[0];
        Assert.Equal(axis == SheetAxis.Row ? "C5:D6" : "E3:F4", Assert.Single(current.MergedCells));
        Assert.Equal(axis == SheetAxis.Row ? 4 : 2, current.Images[0].RowIndex);
        Assert.Equal(axis == SheetAxis.Column ? 4 : 2, current.Images[0].ColumnIndex);
        Assert.Equal(image.Data, current.Images[0].Data);
        session.Undo(); Assert.Equal("C3:D4", Assert.Single(session.Snapshot.Sheets[0].MergedCells));
        session.Redo(); Assert.Equal(current.Images[0].RowIndex, session.Snapshot.Sheets[0].Images[0].RowIndex);
    }

    [Fact]
    public void InsertionInsideMergeExpandsAndOwnerDeletionIsRejected()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Art", mergedCells: ["A2:B4"], images: [new SheetImage([1], "png", 1, 3, 3, 1)]);
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 2, 2));
        Assert.Equal("A2:B6", Assert.Single(session.Snapshot.Sheets[0].MergedCells));
        Assert.Equal(5, session.Snapshot.Sheets[0].Images[0].RowSpan);
        var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1, true)));
        Assert.Same(before, session.Snapshot);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 5, true));
        Assert.Empty(session.Snapshot.Sheets[0].MergedCells); Assert.Empty(session.Snapshot.Sheets[0].Images);
    }

    [Fact]
    public void SplitMergeMovementRejectsAtomicallyAndWholeMovementWorks()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Art", mergedCells: ["A2:B4"]);
        using var session = new DocumentSession(Book(sheet)); var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new MoveAxisCommand(sheet.Id, SheetAxis.Row, 2, 1, 8)));
        Assert.Same(before, session.Snapshot);
        session.Execute(new MoveAxisCommand(sheet.Id, SheetAxis.Row, 1, 3, 8));
        Assert.Equal("A9:B11", Assert.Single(session.Snapshot.Sheets[0].MergedCells));
    }

    [Fact]
    public void MergeCannotHideDataAndNonAnchorWritesAreAtomic()
    {
        var sheet = Sheet("Data", (0, 0, T("title")), (1, 0, T("keep")));
        using var session = new DocumentSession(Book(sheet));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new MergeCellsCommand(sheet.Id, new(0, 0, 1, 1))));
        session.Execute(new MergeCellsCommand(sheet.Id, new(0, 0, 0, 2)));
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([
            new(new(sheet.Id, 2, 0), CellInput.Parse("should not commit")), new(new(sheet.Id, 0, 1), CellInput.Parse("hidden"))])));
        Assert.Same(before, session.Snapshot);
        session.Execute(new MergeCellsCommand(sheet.Id, new(0, 1, 0, 1), true));
        Assert.Empty(session.Snapshot.Sheets[0].MergedCells);
        session.Undo(); Assert.Equal("A1:C1", Assert.Single(session.Snapshot.Sheets[0].MergedCells));
    }

    [Fact]
    public void SettingsRoundtripAndCorruptionFallsBackToBottom()
    {
        var folder = Path.Combine(Path.GetTempPath(), "QuickExcel-settings-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var path = Path.Combine(folder, "settings.json"); var store = new DesktopSettings(path);
            store.Save(new(SheetNavigationPlacement.Right, 150)); Assert.Null(store.Error);
            var reopened = new DesktopSettings(path); reopened.Load(); Assert.Equal(store.Value, reopened.Value);
            File.WriteAllText(path, "{\"Sheets\":99,\"Zoom\":999}"); reopened.Load();
            Assert.Equal(new DesktopPreferences(), reopened.Value); Assert.NotNull(reopened.Error);
            File.WriteAllText(path, "["); reopened.Load(); Assert.Equal(SheetNavigationPlacement.Bottom, reopened.Value.Sheets);
        }
        finally { Directory.Delete(folder, true); }
    }
}
