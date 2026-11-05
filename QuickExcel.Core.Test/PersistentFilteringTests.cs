using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Calculation;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class PersistentFilteringTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "QuickExcel-filter-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static WorkbookServices Services => new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts()));
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public PersistentFilteringTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string FilePath(string name) => Path.Combine(_directory, name);
    private static WorkbookSnapshot Fixture()
    {
        var sheet = Sheet("Filter's", (0, 0, T("Group")), (0, 1, T("Amount")), (1, 0, T("A")), (1, 1, N(30)),
            (2, 0, T("B")), (2, 1, N(10)), (3, 0, T("A")), (3, 1, N(20)), (4, 0, T("A")), (4, 1, N(40)),
            (5, 0, T("outside")), (5, 1, N(100)));
        return Book(new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells, rows: [new(2, new(Hidden: true)), new(3, new(Hidden: true)), new(4, new(Height: 27))]));
    }
    private static WorksheetAutoFilter Filter() => new(new(0, 0, 4, 1), [new(0, new(values: ["A"])),
        new(1, new(first: new(AutoFilterComparison.GreaterThanOrEqual, "20"), second: new(AutoFilterComparison.LessThan, "40"), useAnd: true))]);
    private static void EditPart(string path, Action<XElement> edit)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update); var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
        XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
        edit(xml.Root!); entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
    }

    [Fact]
    public void UntrustedGeneratedRowHintsCannotRemoveSourceRowFormatting()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new SetAutoFilterCommand(id, Filter())); var path = FilePath("row-hints.xlsx"); _storage.Save(path, session.Snapshot);
        EditPart(path, root => root.Descendants(XName.Get("rows", "urn:quickexcel:filter-rows:v1")).Single().SetAttributeValue("generated", "5"));
        var read = _storage.Open(path);
        Assert.Equal(27d, read.Sheets[0].Rows[4].Height);
        Assert.Contains(4, WorksheetFilterEngine.Evaluate(read, read.Sheets[0].Id));
    }

    [Fact]
    public void ImportedFilterUsesLabeledSourceCachesUntilAnEditRequiresFreshResults()
    {
        var sheet = Sheet("Cache", (0, 0, T("Header")), (1, 0, F("=20", 30)), (2, 0, N(10)))
            .WithData(new(new(new(0, 0, 2, 0), [new(0, new(first: new(AutoFilterComparison.GreaterThan, "25")))])));
        var original = Book(sheet);
        var imported = new WorkbookSnapshot(original.DocumentId, original.Revision, original.DateSystem, original.Sheets,
            fileState: new(FilePath("source.xlsx"), 2, [], original, [], []));
        var view = new SheetViewState(); view.Synchronize(imported, sheet.Id);
        Assert.True(view.IsFilteredOut(imported, 0, 2)); Assert.False(view.IsFilteredOut(imported, 0, 1));
        Assert.True(view.FilterUsesSourceCaches); Assert.Null(view.FilterError);
        var statistics = SelectionStatistics.Compute(imported, sheet.Id, [new(1, 0, 2, 0)]);
        Assert.Equal(30d, statistics.Sum); Assert.Equal(1, statistics.OldCaches);
        using var session = new DocumentSession(imported);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 2, 0), CellContent.FromLiteral(CellValue.FromNumber(40)))]));
        view.Synchronize(session.Snapshot, sheet.Id); view.RefreshFilter(session.Snapshot);
        Assert.NotNull(view.FilterError); Assert.False(view.FilterUsesSourceCaches);
        Assert.Throws<InvalidOperationException>(() => SelectionStatistics.Compute(session.Snapshot, sheet.Id, [new(1, 0, 2, 0)]));
        session.Undo(); view.Synchronize(session.Snapshot, sheet.Id); view.RefreshFilter(session.Snapshot);
        // Undo advances the revision; it cannot silently promote a stale imported cache to a fresh result.
        Assert.NotNull(view.FilterError);
    }

    [Fact]
    public void EqualityCanonicalizationAndUnrepresentableFiltersFollowExcelReapplyEvidence()
    {
        var equality = new AutoFilterColumn(first: new(AutoFilterComparison.Equal, "10"));
        Assert.Equal(new[] { "10" }, equality.Values!.Value); Assert.Null(equality.First);
        var matches = WorksheetFilterEngine.Compile(equality);
        Assert.True(matches(CellValue.FromText("10"), new(), ExcelDateSystem.Windows1900));
        Assert.True(matches(CellValue.FromNumber(10), new(), ExcelDateSystem.Windows1900));
        Assert.False(matches(CellValue.FromNumber(10), new() { NumberFormatCode = "0.00" }, ExcelDateSystem.Windows1900));
        var tilde = new AutoFilterColumn(first: new(AutoFilterComparison.Equal, "a~~b")); Assert.Equal("a~~b", tilde.Values!.Value.Single());
        Assert.Equal("a*b", new AutoFilterColumn(first: new(AutoFilterComparison.Equal, "a~*b")).Values!.Value.Single());
        Assert.True(new AutoFilterColumn(first: new(AutoFilterComparison.Equal, "")).IncludeBlank);
        var error = WorksheetFilterEngine.Compile(new(first: new(AutoFilterComparison.NotEqual, "#N/A")));
        Assert.True(error(CellValue.FromText("#N/A"), new(), ExcelDateSystem.Windows1900));
        Assert.False(error(CellValue.FromError(CellErrorCode.NotAvailable), new(), ExcelDateSystem.Windows1900));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(values: []));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(first: new(AutoFilterComparison.Equal, " ")));
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.ValueLabel(CellValue.FromNumber(1), new() { NumberFormatCode = "[DBNum1]0" }, ExcelDateSystem.Windows1900));
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.ValueLabel(CellValue.FromNumber(1), new() { NumberFormatCode = "_(*0.00)" }, ExcelDateSystem.Windows1900));
    }

    [Fact]
    public void FilterHistoryBudgetAndExplicitStaleValueExportRemainAtomic()
    {
        var source = Fixture(); using var limited = new DocumentSession(source, historyBudgetBytes: 2000);
        Assert.Throws<InvalidOperationException>(() => limited.Execute(new SetAutoFilterCommand(source.Sheets[0].Id, new(new(0, 0, 4, 1), [new(0, new(values: [new string('x', 2000)]))]))));
        Assert.Same(source, limited.Snapshot); Assert.False(limited.CanUndo);
        var sheet = Sheet("Cache", (0, 0, T("Header")), (1, 0, F("=20", 30))).WithData(new(new(new(0, 0, 1, 0), [new(0, new(first: new(AutoFilterComparison.GreaterThan, "25")))])));
        var book = Book(sheet); var path = FilePath("values.xlsx");
        Assert.Throws<InvalidOperationException>(() => _storage.Export(path, book));
        _storage.Export(path, book, allowStaleCaches: true, discardUnsupported: false);
        var read = _storage.Open(path); Assert.Equal(CellContentKind.Literal, read.Sheets[0].GetCell(1, 0)!.Content.Kind);
        Assert.True(sheet.Data.ContentEquals(read.Sheets[0].Data)); Assert.Empty(WorksheetFilterEngine.Evaluate(read, read.Sheets[0].Id));
    }

    [Fact]
    public void NativeExcelFilterAndSortDefinitionsMatchVisibleRowsAndRoundTrip()
    {
        using var spec = ScalarFormulaFixture.Load("filter-persistence-spec"); using var oracle = ScalarFormulaFixture.Load("excel-filter-persistence");
        var inputs = spec.RootElement.GetProperty("Rows").EnumerateArray().ToArray();
        foreach (var expected in oracle.RootElement.GetProperty("Cases").EnumerateArray())
        {
            var name = expected.GetProperty("Id").GetString()!; var styles = new List<StyleDefinition> { new() }; var cells = new Dictionary<CellPosition, CellSnapshot>();
            cells[new(0, 0)] = T("Label"); cells[new(0, 1)] = T("Group"); cells[new(0, 2)] = T("Identity");
            var order = expected.GetProperty("Order").EnumerateArray().Select(v => (int)v.GetDouble()).ToArray();
            for (var r = 0; r < order.Length; r++)
            {
                var input = inputs[order[r] - 1]; var kind = input.GetProperty("Kind").GetString(); var style = 0;
                if (input.TryGetProperty("Format", out var format)) { style = styles.Count; styles.Add(new() { NumberFormatCode = format.GetString() }); }
                var value = kind switch { "Number" => CellValue.FromNumber(input.GetProperty("Value").GetDouble()), "Text" => CellValue.FromText(input.GetProperty("Value").GetString()!), _ => CellValue.Blank };
                CellSnapshot cell;
                if (kind == "Formula")
                {
                    var formula = input.GetProperty("Value").GetString()!;
                    value = formula switch { "=TRUE" => CellValue.FromBoolean(true), "=FALSE" => CellValue.FromBoolean(false), "=#N/A" => CellValue.FromError(CellErrorCode.NotAvailable), "=#VALUE!" => CellValue.FromError(CellErrorCode.Value), _ => CellValue.FromText("") };
                    cell = new(CellContent.FromFormula(formula, style), new(value, value.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh, 0, 0, 0));
                }
                else cell = new(CellContent.FromLiteral(value, style));
                cells[new(r + 1, 0)] = cell; cells[new(r + 1, 1)] = N((order[r] - 1) % 6 + 1); cells[new(r + 1, 2)] = N(order[r]);
            }
            var expectedHidden = expected.GetProperty("HiddenRows").EnumerateArray().Select(v => v.GetInt32() - 1).ToArray();
            var source = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new(SheetId.New(), name, cells, rows: expectedHidden.Select(r => new KeyValuePair<int, RowMetadata>(r, new(Hidden: true))))], styles: styles);
            var path = FilePath(name + ".xlsx"); _storage.Save(path, source);
            EditPart(path, root =>
            {
                var data = root.Element(Ns + "sheetData")!;
                foreach (var key in new[] { "sortStateXml", "autoFilterXml" })
                    if (expected.GetProperty(key).GetString() is { } text) data.AddAfterSelf(XElement.Parse(text));
            });
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/styles.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
                xml.Root!.Add(XElement.Parse(oracle.RootElement.GetProperty("DifferentialStylesXml").GetString()!));
                entry.Delete(); using var streamOut = zip.CreateEntry("xl/styles.xml").Open(); xml.Save(streamOut);
            }
            using var document = new WorkbookDocument(_storage.Open(path), Services); document.Calculate();
            var sheet = document.Snapshot.Sheets[0]; Assert.NotNull(sheet.Data.Filter); Assert.Empty(document.Snapshot.FileState!.CompatibilityIssues);
            if (name.StartsWith("sort-", StringComparison.Ordinal)) Assert.NotNull(sheet.Data.Sort);
            Assert.Equal(expectedHidden.Order(), WorksheetFilterEngine.Evaluate(document.Snapshot, sheet.Id).Union(sheet.Rows.Where(p => p.Value.Hidden).Select(p => p.Key)).Order());
            document.SetCell(0, 40, 3, "outside edit"); document.Calculate();
            foreach (var extension in new[] { ".xlsx", ".xlsq" })
            {
                var output = FilePath(name + "-output" + extension); _storage.Save(output, document.Snapshot); var reopened = _storage.Open(output);
                Assert.True(sheet.Data.ContentEquals(reopened.Sheets[0].Data)); Assert.Empty(reopened.FileState!.CompatibilityIssues);
            }
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void MulticolumnFiltersAndManualVisibilitySurviveClearUndoAndOrdinarySave(string extension)
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        var original = session.Snapshot;
        session.Execute(new SetAutoFilterCommand(id, Filter()));
        Assert.True(session.IsDirty); Assert.Equal(new[] { 2, 4 }, WorksheetFilterEngine.Evaluate(session.Snapshot, id).Order());
        var saved = session.Snapshot; var path = FilePath("filtered" + extension); _storage.Save(path, saved);
        var reopened = _storage.Open(path); var sheet = reopened.Sheets[0];
        Assert.Empty(reopened.FileState!.CompatibilityIssues); Assert.True(saved.Sheets[0].Data.ContentEquals(sheet.Data));
        Assert.Equal(saved.Sheets[0].Rows.OrderBy(p => p.Key), sheet.Rows.OrderBy(p => p.Key));
        Assert.True(sheet.Rows[2].Hidden); Assert.True(sheet.Rows[3].Hidden); Assert.False(sheet.Rows[4].Hidden);
        using var opened = new DocumentSession(reopened);
        opened.Execute(new SetAutoFilterCommand(sheet.Id, null)); Assert.Empty(WorksheetFilterEngine.Evaluate(opened.Snapshot, sheet.Id));
        Assert.True(opened.Snapshot.Sheets[0].Rows[2].Hidden); Assert.True(opened.Snapshot.Sheets[0].Rows[3].Hidden);
        opened.Undo(); Assert.True(saved.Sheets[0].Data.ContentEquals(opened.Snapshot.Sheets[0].Data)); opened.Redo();
        _storage.Save(FilePath("cleared" + extension), opened.Snapshot);
        session.Undo(); Assert.True(original.Sheets[0].Data.ContentEquals(session.Snapshot.Sheets[0].Data)); session.Redo();
        Assert.True(saved.Sheets[0].Data.ContentEquals(session.Snapshot.Sheets[0].Data));
    }

    [Fact]
    public void SortStateIncludesNoOpSortAndFollowsAxesCopyRenameAndUndo()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new SetAutoFilterCommand(id, Filter()));
        session.Execute(new SortRangeCommand(new(id, new(0, 0), new(4, 1), [new(0), new(1, true)], hasHeader: true)));
        var data = session.Snapshot.Sheets[0].Data; Assert.Equal(new CellRectangle(1, 0, 4, 1), data.Sort!.Range);
        Assert.Equal(40, session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue);
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 2, 1));
        Assert.Equal(5, session.Snapshot.Sheets[0].Data.Filter!.Range.Bottom); Assert.Equal(5, session.Snapshot.Sheets[0].Data.Sort!.Range.Bottom);
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Column, 0, 1)); Assert.Equal(new[] { 1, 2 }, session.Snapshot.Sheets[0].Data.Sort!.Keys.Select(k => k.Column));
        session.Undo(); session.Undo(); Assert.True(data.ContentEquals(session.Snapshot.Sheets[0].Data));
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1, delete: true)));
        session.Execute(new RenameSheetCommand(id, "New name"));
        var copy = new CopySheetCommand(id, "Copied", 1); session.Execute(copy);
        Assert.True(data.ContentEquals(session.Snapshot.GetSheet(copy.Id).Data)); session.Undo();
        var path = FilePath("sorted.xlsx"); _storage.Save(path, session.Snapshot); Assert.True(data.ContentEquals(_storage.Open(path).Sheets[0].Data));
        using var noOp = new DocumentSession(Book(Sheet("S", (0, 0, N(1)), (1, 0, N(2)))));
        var sameId = noOp.Snapshot.Sheets[0].Id;
        Assert.True(noOp.Execute(new SortRangeCommand(sameId, new(0, 0), new(1, 0), 0)).Changed); Assert.NotNull(noOp.Snapshot.Sheets[0].Data.Sort);
        var sorted = noOp.Snapshot.Sheets[0];
        noOp.Execute(new ClearSortStateCommand(sameId)); Assert.Null(noOp.Snapshot.Sheets[0].Data.Sort);
        Assert.Same(sorted.Cells, noOp.Snapshot.Sheets[0].Cells); noOp.Undo(); Assert.NotNull(noOp.Snapshot.Sheets[0].Data.Sort);
    }

    [Theory]
    [InlineData(SortBasis.FillColor)]
    [InlineData(SortBasis.FontColor)]
    public void ColorAndCustomSortStateSerializeDifferentialStyleRelationships(SortBasis basis)
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new SortRangeCommand(new(id, new(0, 0), new(4, 1), [new(0, basis: basis, preferredColor: "ff12ab34"), new(1)], true)));
        var path = FilePath("color.xlsx"); _storage.Save(path, session.Snapshot);
        var read = _storage.Open(path); Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(session.Snapshot.Sheets[0].Data.ContentEquals(read.Sheets[0].Data));
        session.Execute(new SortRangeCommand(new(id, new(0, 0), new(4, 1), [new(0, customOrder: ["B", "A"])], true)));
        _storage.Save(FilePath("custom.xlsx"), session.Snapshot);
        Assert.Equal(new[] { "B", "A" }, _storage.Open(FilePath("custom.xlsx")).Sheets[0].Data.Sort!.Keys[0].CustomOrder);
    }

    [Fact]
    public void UnsupportedFilterStaysInOriginalPackageAndModifiedNormalSaveIsBlocked()
    {
        var path = FilePath("unknown.xlsx"); _storage.Save(path, Fixture());
        EditPart(path, root => root.Element(Ns + "sheetData")!.AddAfterSelf(new XElement(Ns + "autoFilter", new XAttribute("ref", "A1:B5"),
            new XElement(Ns + "filterColumn", new XAttribute("colId", 1), new XElement(Ns + "iconFilter", new XAttribute("iconSet", "3Arrows"), new XAttribute("iconId", 2))))));
        var bytes = File.ReadAllBytes(path); var read = _storage.Open(path); Assert.Null(read.Sheets[0].Data.Filter);
        Assert.Contains(read.FileState!.CompatibilityDiagnostics, d => d.Feature == "worksheet/autoFilter");
        _storage.Save(FilePath("unchanged.xlsx"), read); Assert.Equal(bytes, File.ReadAllBytes(FilePath("unchanged.xlsx")));
        using var session = new DocumentSession(read); session.Execute(new SetCellsCommand([new(new(read.Sheets[0].Id, 6, 0), CellContent.FromLiteral(CellValue.FromNumber(7)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(path, session.Snapshot)); Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public void FilterEvaluationReappliesEditsAndRefusesStaleResultsWithoutChangingManualRows()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new SetAutoFilterCommand(id, Filter()));
        session.Execute(new SetCellsCommand([new(new(id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(50)))]));
        Assert.Contains(1, WorksheetFilterEngine.Evaluate(session.Snapshot, id));
        session.Execute(new SetCellsCommand([new(new(id, 1, 1), CellContent.FromFormula("=30"))]));
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.Evaluate(session.Snapshot, id));
        Assert.Throws<InvalidOperationException>(() => _storage.Save(FilePath("stale.xlsx"), session.Snapshot));
        Assert.True(session.Snapshot.Sheets[0].Rows[2].Hidden);
        Assert.Throws<ArgumentException>(() => new WorksheetAutoFilter(new(0, 0, 99_999, 25), Enumerable.Range(0, 26).Select(i => new KeyValuePair<int, AutoFilterColumn>(i, new(values: ["x"])) )));
        var source = Fixture(); var filtered = source.WithSheets([source.Sheets[0].WithData(new(Filter()))]);
        Assert.Throws<OperationCanceledException>(() => WorksheetFilterEngine.Evaluate(filtered, source.Sheets[0].Id, cancellationToken: new(true)));
    }
}
