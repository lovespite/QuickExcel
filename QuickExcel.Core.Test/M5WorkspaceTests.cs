using System.Globalization;
using QuickExcel.Calculation;
using QuickExcel.Desktop;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M5WorkspaceTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-M5-workspace-" + Guid.NewGuid().ToString("N"));
    public M5WorkspaceTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private string FilePath(string name = "book.xlsq") => Path.Combine(_folder, name);
    private static WorkbookServices Services => new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts()));
    private static CellSnapshot B(bool value) => new(CellContent.FromLiteral(CellValue.FromBoolean(value)));
    private static CellSnapshot E(CellErrorCode value) => new(CellContent.FromLiteral(CellValue.FromError(value)));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SortIsStableTypedBlankLastAndUndoable(bool descending)
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (0, 1, T("first")), (1, 0, N(1)), (1, 1, T("small")),
            (2, 0, N(2)), (2, 1, T("second")), (3, 1, T("blank")), (4, 0, T("001")), (5, 0, B(true)), (6, 0, E(CellErrorCode.Value)));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(6, 1), 0, descending));
        var sorted = session.Snapshot.Sheets[0];
        Assert.Equal("blank", sorted.GetCell(6, 1)!.Value.TextValue);
        var labels = Enumerable.Range(0, 7).Select(r => sorted.GetCell(r, 1)?.Value).ToArray();
        Assert.True(Array.IndexOf(labels, CellValue.FromText("first")) < Array.IndexOf(labels, CellValue.FromText("second")));
        Assert.Equal(descending ? CellValueKind.Error : CellValueKind.Number, sorted.GetCell(0, 0)!.Value.Kind);
        Assert.True(session.Undo()); Assert.Equal(sheet.Cells, session.Snapshot.Sheets[0].Cells);
        Assert.True(session.Redo()); Assert.Equal(sorted.Cells, session.Snapshot.Sheets[0].Cells);
    }

    [Fact]
    public void SortMovesRelativeMixedCrossSheetFormulasAndLeavesOutsideReferencesAndNamesAtAddresses()
    {
        var data = Sheet("Data", (1, 0, N(2)), (1, 1, F("=A2+$A2+A$2+$A$2+Other!A2")), (2, 0, N(1)), (2, 1, F("=A3")), (9, 0, F("=A2")));
        var other = Sheet("Other", (0, 0, F("=Data!A2")), (1, 0, N(3)));
        using var session = new DocumentSession(Book(data, other).WithDefinedNames([new("Pinned", "=Data!$A$2")]));
        session.Execute(new SortRangeCommand(data.Id, new(1, 0), new(2, 1), 0));
        var moved = session.Snapshot.Sheets[0].GetCell(2, 1)!.Content.FormulaText;
        var refs = FormulaSyntax.Walk(FormulaSyntax.Parse(moved)).OfType<ReferenceNode>().ToArray();
        Assert.Equal(new[] { 2, 2, 1, 1, 2 }, refs.Select(r => r.Row));
        Assert.Equal("Other", refs[4].Sheet);
        Assert.Equal("=A2", session.Snapshot.Sheets[0].GetCell(9, 0)!.Content.FormulaText);
        Assert.Equal("=Data!A2", session.Snapshot.Sheets[1].GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal("=Data!$A$2", session.Snapshot.DefinedNames[0].Formula);
    }

    [Fact]
    public void SortBoundaryReferenceBecomesRefAndAbsoluteSingleReferenceDoesNotMoveItsEnd()
    {
        var sheet = Sheet("D", (0, 0, N(2)), (0, 1, F("=$A$100000")), (1, 0, N(1)), (1, 1, F("=A1")));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 1), 0));
        Assert.Equal("=#REF!", session.Snapshot.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        Assert.Equal("=$A$100000", session.Snapshot.Sheets[0].GetCell(1, 1)!.Content.FormulaText);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SortRejectsMergedOrImageSheetsAtomically(bool merge)
    {
        var sheet = Sheet("D", (0, 0, N(2)), (1, 0, N(1)));
        sheet = new(sheet.Id, sheet.Name, sheet.Cells, mergedCells: merge ? ["A1:B1"] : [],
            images: merge ? [] : [new SheetImage([1], "png", 0, 0)]);
        using var session = new DocumentSession(Book(sheet));
        Assert.Throws<NotSupportedException>(() => session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0)));
        Assert.Equal(0, session.Revision); Assert.False(session.CanUndo);
    }

    [Fact]
    public void SortRejectsStaleKeysAndCapacityButFreshKeysWork()
    {
        var sheet = Sheet("D", (0, 0, F("=2")), (1, 0, F("=1")));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        Assert.Throws<InvalidOperationException>(() => doc.Session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0)));
        doc.Calculate();
        Assert.True(doc.Session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0)).Changed);
        Assert.Throws<ArgumentException>(() => doc.Session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(20000, 0), 0)));
    }

    [Fact]
    public void SortNoOpCreatesNoHistoryAndPreservesRowsOutsideRectangle()
    {
        var sheet = Sheet("D", (0, 0, N(1)), (1, 0, N(2)), (0, 1, T("outside")));
        using var session = new DocumentSession(Book(sheet));
        Assert.False(session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0)).Changed);
        Assert.False(session.CanUndo);
        session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0, true));
        Assert.Equal("outside", session.Snapshot.Sheets[0].GetCell(0, 1)!.Value.TextValue);
    }

    [Fact]
    public void StylesApplyAtomicallyPreserveRawInputDeduplicateAndRoundTrip()
    {
        using var doc = new WorkbookDocument(Book(Sheet("D", (0, 0, N(0.125)), (1, 0, F("=A1*2")))), Services);
        var sheet = doc.Snapshot.Sheets[0];
        var patch = new CellStylePatch(Bold: true, Italic: true, FillColor: "FFFFE599", Horizontal: 2, NumberFormatId: 10, Border: true);
        doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0), new(sheet.Id, 1, 0), new(sheet.Id, 2, 0)], patch));
        Assert.Equal(2, doc.Snapshot.Styles.Length);
        Assert.Equal("0.125", doc.GetCellInput(0, 0, 0));
        Assert.Equal("12.50%", CellDisplayFormatter.Format(doc.Snapshot.Sheets[0].GetCell(0, 0)!.Value, doc.Snapshot.Styles[1], ExcelDateSystem.Windows1900, CultureInfo.InvariantCulture));
        Assert.Equal("=A1*2", doc.GetCellInput(0, 1, 0));
        doc.Calculate(); doc.Save(FilePath());
        using var reopened = WorkbookDocument.Open(FilePath(), Services);
        Assert.Equal(doc.Snapshot.Styles.ToArray(), reopened.Snapshot.Styles.ToArray());
        Assert.Equal(CellValue.Blank, reopened.Snapshot.Sheets[0].GetCell(2, 0)!.Value);
        Assert.True(doc.Undo()); Assert.Single(doc.Snapshot.Styles);
        Assert.True(doc.Redo()); Assert.Equal(2, doc.Snapshot.Styles.Length);
        var before = doc.Snapshot;
        Assert.Throws<ArgumentException>(() => doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(FontColor: "invalid"))));
        Assert.Same(before, doc.Snapshot);
    }

    [Fact]
    public void RibbonFontAndAlignmentPatchPreservesContentAndRoundTripsWithUndo()
    {
        using var doc = new WorkbookDocument(Book(Sheet("D", (0, 0, F("=1+2")), (1, 0, T("long text")))), Services);
        var id = doc.Snapshot.Sheets[0].Id;
        var addresses = new[] { new CellAddress(id, 0, 0), new CellAddress(id, 1, 0) };
        doc.Session.Execute(new SetCellStyleCommand(addresses, new(FontName: "Consolas", FontSize: 18, WrapText: true, Vertical: 1)));
        Assert.Equal("=1+2", doc.GetCellInput(0, 0, 0));
        Assert.Equal("long text", doc.GetCellInput(0, 1, 0));
        Assert.Equal(new StyleDefinition { FontName = "Consolas", FontSize = 18, WrapText = true, Vertical = 1 }, doc.Snapshot.Styles[1]);
        Assert.Equal(doc.Snapshot.Sheets[0].GetStyleId(0, 0), doc.Snapshot.Sheets[0].GetStyleId(1, 0));
        doc.Save(FilePath());
        using var reopened = WorkbookDocument.Open(FilePath(), Services);
        Assert.Equal(doc.Snapshot.Styles.ToArray(), reopened.Snapshot.Styles.ToArray());
        Assert.True(doc.Undo()); Assert.Single(doc.Snapshot.Styles);
        Assert.True(doc.Redo()); Assert.Equal(reopened.Snapshot.Styles[1], doc.Snapshot.Styles[1]);
        var before = doc.Snapshot;
        foreach (var invalid in new[] { new CellStylePatch(FontName: " "), new(FontSize: double.NaN), new(FontSize: 0), new(Vertical: 4) })
        {
            Assert.Throws<ArgumentException>(() => doc.Session.Execute(new SetCellStyleCommand(addresses, invalid)));
            Assert.Same(before, doc.Snapshot);
        }
    }

    [Fact]
    public void ResetBlankStyleOverridesInheritedRowStyleAndSurvivesSave()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "D", rows: [new(0, new(StyleId: 1))]);
        using var doc = new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Bold = true }]), Services);
        doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(Reset: true)));
        Assert.Equal(0, doc.Snapshot.Sheets[0].GetStyleId(0, 0));
        doc.Save(FilePath());
        using var reopened = WorkbookDocument.Open(FilePath(), Services);
        Assert.Equal(0, reopened.Snapshot.Sheets[0].GetStyleId(0, 0));
    }

    [Theory]
    [InlineData(SheetAxis.Row)]
    [InlineData(SheetAxis.Column)]
    public void DimensionsAreUndoableAndPersisted(SheetAxis axis)
    {
        using var doc = WorkbookDocument.CreateBlank(Services); var id = doc.Snapshot.Sheets[0].Id;
        doc.Session.Execute(new SetAxisMetadataCommand(id, axis, 1, 2, 25, true));
        doc.Save(FilePath()); using var reopened = WorkbookDocument.Open(FilePath(), Services);
        if (axis == SheetAxis.Row) Assert.Equal(new RowMetadata(25, true), reopened.Snapshot.Sheets[0].Rows[1]);
        else Assert.Equal(new ColumnMetadata(25, true), reopened.Snapshot.Sheets[0].Columns[1]);
        Assert.True(doc.Undo()); Assert.Empty(doc.Snapshot.Sheets[0].Rows); Assert.Empty(doc.Snapshot.Sheets[0].Columns);
    }

    [Fact]
    public void DeleteSheetUpdatesReferencesRemovesScopedNamesAndUndoRestores()
    {
        var data = Sheet("Data", (0, 0, N(42))); var report = Sheet("R", (0, 0, F("=Data!A1")));
        using var doc = new WorkbookDocument(Book(data, report).WithDefinedNames([new("Local", "=Data!$A$1", data.Id), new("Global", "=Data!$A$1")]), Services);
        doc.Session.Execute(new DeleteSheetCommand(data.Id));
        Assert.Single(doc.Snapshot.Sheets); Assert.Single(doc.Snapshot.DefinedNames);
        doc.Calculate(); Assert.Equal(CellValue.FromError(CellErrorCode.Reference), doc.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.True(doc.Undo()); Assert.Equal(2, doc.Snapshot.DefinedNames.Length); Assert.Equal(data.Cells, doc.Snapshot.Sheets[0].Cells);
        Assert.True(doc.Redo());
        Assert.Throws<InvalidOperationException>(() => doc.Session.Execute(new DeleteSheetCommand(report.Id)));
    }

    [Fact]
    public void FilterUsesDisplayTextPreservesHeaderAndModelAndCopySkipsOnlyFilteredRows()
    {
        var sheet = Sheet("D", (0, 0, T("Header")), (1, 0, T("keep")), (2, 0, T("drop")), (3, 0, T("KEEP")));
        var snapshot = Book(sheet); var view = new SheetViewState { FilterRange = new(0, 0, 3, 0), FilterColumn = 0, FilterText = "Keep" };
        bool Excluded(int row) => view.IsFilteredOut(snapshot, 0, row);
        Assert.False(Excluded(0)); Assert.False(Excluded(1)); Assert.True(Excluded(2)); Assert.False(Excluded(3)); Assert.False(Excluded(4));
        var rectangle = GridRange.Rectangle([new(0, 0), new(1, 0), new(3, 0)], Excluded);
        Assert.Equal("Header\r\nkeep\r\nKEEP", GridClipboard.Write(rectangle, (r, c) => sheet.GetCell(r, c)!.Value.TextValue, Excluded));
        Assert.Throws<InvalidOperationException>(() => view.RequireVisibleTargets(snapshot, 0, [1, 2, 3]));
        view.RequireVisibleTargets(snapshot, 0, [1, 3]);
        Assert.Equal(4, sheet.Cells.Count);
        Assert.Throws<ArgumentException>(() => GridRange.Rectangle([new(0, 0), new(3, 0)], Excluded));
    }

    [Fact]
    public void RecentFilesAreBoundedDeduplicatedPersistentAndTolerateMalformedSettings()
    {
        var config = FilePath("recent.json"); var recent = new RecentFileStore(config);
        for (var i = 0; i < 12; i++) recent.Remember(FilePath(i + ".xlsq"));
        recent.Remember(FilePath("11.xlsq").ToUpperInvariant());
        var loaded = new RecentFileStore(config); loaded.Load();
        Assert.Equal(10, loaded.Paths.Count); Assert.Equal(FilePath("11.xlsq").ToUpperInvariant(), loaded.Paths[0]);
        loaded.Clear(); recent.Load(); Assert.Empty(recent.Paths);
        File.WriteAllText(config, "{broken"); recent.Load(); Assert.Empty(recent.Paths); Assert.NotNull(recent.Error);
    }

    [Theory]
    [InlineData("replace")]
    [InlineData("delete")]
    [InlineData("same-length-and-time")]
    public void ExternalConflictNeverAdvancesSavedRevisionOrOverwritesExternalVersion(string mode)
    {
        using var doc = WorkbookDocument.CreateBlank(Services);
        var path = FilePath(); var version = doc.SaveChecked(path, new(null));
        doc.SetCell(0, 0, 0, "unsaved");
        if (mode == "delete") File.Delete(path);
        else if (mode == "replace") File.WriteAllText(path, "external");
        else
        {
            var timestamp = File.GetLastWriteTimeUtc(path); var bytes = File.ReadAllBytes(path); bytes[0] ^= 1;
            File.WriteAllBytes(path, bytes); File.SetLastWriteTimeUtc(path, timestamp);
        }
        var external = File.Exists(path) ? File.ReadAllBytes(path) : null;
        Assert.Throws<WorkbookFileConflictException>(() => doc.SaveChecked(path, version));
        Assert.True(doc.IsDirty); Assert.Equal(0, doc.Session.SavedRevision);
        Assert.Equal(external, File.Exists(path) ? File.ReadAllBytes(path) : null);
        Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
        doc.SaveChecked(FilePath("copy.xlsq"), new(null)); Assert.False(doc.IsDirty);
    }

    [Fact]
    public void NewTargetCreatedBySomeoneElseIsAConflictAndFailedSaveKeepsDirty()
    {
        using var doc = WorkbookDocument.CreateBlank(Services); doc.SetCell(0, 0, 0, "draft");
        File.WriteAllText(FilePath(), "external");
        Assert.Throws<WorkbookFileConflictException>(() => doc.SaveChecked(FilePath(), new(null)));
        Assert.True(doc.IsDirty);
        Assert.Throws<DirectoryNotFoundException>(() => doc.SaveChecked(FilePath("missing/book.xlsq"), new(null)));
        Assert.True(doc.IsDirty);
    }

    [Theory]
    [InlineData(2, "12.35")]
    [InlineData(10, "1234.50%")]
    public void NumericDisplayNeverChangesUnderlyingScalar(int format, string expected)
    {
        var value = CellValue.FromNumber(12.345);
        Assert.Equal(expected, CellDisplayFormatter.Format(value, new() { NumberFormatId = format }, ExcelDateSystem.Windows1900, CultureInfo.InvariantCulture));
        Assert.Equal(12.345, value.NumberValue);
    }

    [Fact]
    public void VersionIsRecheckedAfterStagedFileValidationAndConflictLeavesNoTemp()
    {
        var path = FilePath(); File.WriteAllText(path, "original");
        var expected = FileVersions.Read(path);
        var atomic = typeof(TypedWorkbookStorage).Assembly.GetType("QuickExcel.Storage.PackageIO")!
            .GetMethod("AtomicWrite", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)!;
        var error = Assert.Throws<System.Reflection.TargetInvocationException>(() => atomic.Invoke(null,
            [path, new byte[] { 1, 2 }, (Action<string>)(_ => File.WriteAllText(path, "external-during-validation")), CancellationToken.None, expected]));
        Assert.IsType<WorkbookFileConflictException>(error.InnerException);
        Assert.Equal("external-during-validation", File.ReadAllText(path));
        Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
    }

    [Fact]
    public void SortStylesMoveWithCellsButRowDimensionsAndOutsideStylesStayAtTheirAddresses()
    {
        var sheet = Sheet("D", (0, 0, N(2)), (1, 0, N(1)));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(Bold: true)));
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 0, 1, 40));
        session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 0), 0));
        Assert.Equal(1, session.Snapshot.Sheets[0].GetStyleId(1, 0));
        Assert.Equal(0, session.Snapshot.Sheets[0].GetStyleId(0, 0));
        Assert.Equal(40, session.Snapshot.Sheets[0].Rows[0].Height);
    }

    [Fact]
    public void InvalidMovedFormulaAbortsEntireSortBeforeHistoryChanges()
    {
        var sheet = Sheet("D", (0, 0, N(2)), (1, 0, N(1)), (1, 1, F("=NOT VALID(")));
        using var session = new DocumentSession(Book(sheet)); var before = session.Snapshot;
        Assert.Throws<FormatException>(() => session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(1, 1), 0)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void InvalidLastStyleTargetDoesNotCommitEarlierChanges()
    {
        var sheet = Sheet("D", (0, 0, N(1))); using var session = new DocumentSession(Book(sheet)); var before = session.Snapshot;
        Assert.Throws<ArgumentOutOfRangeException>(() => session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0), new(sheet.Id, 100000, 0)], new(Bold: true))));
        Assert.Same(before, session.Snapshot); Assert.Single(session.Snapshot.Styles); Assert.False(session.CanUndo);
    }

    [Fact]
    public void CustomDatePatchPersistsWithNormalizedFormatId()
    {
        using var doc = new WorkbookDocument(Book(Sheet("D", (0, 0, N(46000)))), Services);
        var id = doc.Snapshot.Sheets[0].Id;
        doc.Session.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(NumberFormatId: 0, NumberFormatCode: "yyyy-mm-dd")));
        doc.Save(FilePath()); using var reopened = WorkbookDocument.Open(FilePath(), Services);
        Assert.Equal("2025-12-09", reopened.GetCellValue(0, 0, 0)); Assert.Equal("46000", reopened.GetCellInput(0, 0, 0));
    }

    [Fact]
    public void StyleTableOverflowRejectsBeforeCreatingUnsavableState()
    {
        var sheet = Sheet("D", (0, 0, N(1)));
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet],
            styles: Enumerable.Range(0, 4096).Select(i => new StyleDefinition { FontName = "font" + i })));
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(Bold: true))));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void DeletingImageSheetMustFitUndoBudget()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "D", images: [new SheetImage(new byte[4096], "png", 0, 0)]);
        using var session = new DocumentSession(Book(sheet, Sheet("Other")), historyBudgetBytes: 2048);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new DeleteSheetCommand(sheet.Id)));
        Assert.Equal(2, session.Snapshot.Sheets.Length); Assert.False(session.CanUndo);
    }
}
