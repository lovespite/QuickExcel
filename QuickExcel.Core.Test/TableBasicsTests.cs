using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Desktop;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class TableBasicsTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "QuickExcel-tables-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    public TableBasicsTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string FilePath(string name) => Path.Combine(_directory, name);
    private static WorkbookSnapshot Fixture() => Book(Sheet("Data", (0, 0, T("Group")), (0, 1, T("Amount")), (1, 0, T("B")), (1, 1, N(20)), (2, 0, T("A")), (2, 1, N(10))));

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void CreateResizeSortAndOrdinarySaveRetainTableIdentityAndCellContent(string extension)
    {
        using var session = new DocumentSession(Fixture()); var original = session.Snapshot; var sheet = original.Sheets[0];
        session.Execute(new CreateTableCommand(sheet.Id, new(0, 0, 2, 1), "Orders", TableStyles.Resolve(original, "TableStyleMedium2")));
        Assert.Same(sheet.Cells, session.Snapshot.Sheets[0].Cells); Assert.True(session.IsDirty);
        var table = session.Snapshot.Sheets[0].Data.Tables.Single(); Assert.Equal("FF156082", TablePresentation.GetStyle(session.Snapshot, session.Snapshot.Sheets[0], 0, 0).FillColor);
        session.Execute(new ResizeTableCommand(sheet.Id, table.Id, 4)); Assert.Equal(4, session.Snapshot.Sheets[0].Data.Tables[0].Filter!.Range.Bottom);
        Assert.Same(sheet.Cells, session.Snapshot.Sheets[0].Cells); session.Undo(); Assert.Equal(2, session.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom);
        session.Execute(new SortRangeCommand(new(sheet.Id, new(0, 0), new(2, 1), [new(1)], true)));
        Assert.Equal("A", session.Snapshot.Sheets[0].GetCell(1, 0)!.Value.TextValue); Assert.Null(session.Snapshot.Sheets[0].Data.Sort); Assert.NotNull(session.Snapshot.Sheets[0].Data.Tables[0].Sort);
        var path = FilePath("table." + extension); _storage.Save(path, session.Snapshot); var read = _storage.Open(path);
        Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(session.Snapshot.Sheets[0].Data.ContentEquals(read.Sheets[0].Data));
        Assert.Equal("A", read.Sheets[0].GetCell(1, 0)!.Value.TextValue);
        session.Undo(); Assert.Equal("B", session.Snapshot.Sheets[0].GetCell(1, 0)!.Value.TextValue);
        session.Undo(); Assert.Empty(session.Snapshot.Sheets[0].Data.Tables); Assert.Same(sheet.Cells, session.Snapshot.Sheets[0].Cells);
    }

    [Fact]
    public void AllSixtyBuiltinTableStylesMatchExcelHeaderStripesFirstLastAndTotalColors()
    {
        using var oracle = ScalarFormulaFixture.Load("excel-table-style-display"); var empty = Fixture();
        foreach (var record in oracle.RootElement.GetProperty("After").EnumerateArray())
        {
            var style = TableStyles.Resolve(empty, record.GetProperty("Style").GetString()!, rows: true, columns: true, first: true, last: true);
            var table = new WorksheetTable(1, "DataTable", new(1, 1, 7, 4), [new(1, "A"), new(2, "B"), new(3, "C"), new(4, "D")], totalsRow: true, style: style);
            var sheet = empty.Sheets[0].WithData(new() { Tables = [table] }); var book = Book(sheet);
            foreach (var cell in record.GetProperty("Cells").EnumerateArray())
            {
                var address = CellRectangle.Parse(cell.GetProperty("Address").GetString()!); var actual = TablePresentation.GetStyle(book, sheet, address.Top, address.Left);
                Assert.Equal(cell.GetProperty("Bold").GetBoolean(), actual.Bold); Assert.Equal(cell.GetProperty("Font").GetString(), actual.FontColor ?? "FF000000");
                Assert.Equal(cell.GetProperty("Fill").GetString(), actual.FillColor ?? "FFFFFFFF");
                foreach (var (edge, kind, color) in new[] { ("Left", actual.LeftBorder, actual.LeftBorderColor), ("Right", actual.RightBorder, actual.RightBorderColor),
                    ("Top", actual.TopBorder, actual.TopBorderColor), ("Bottom", actual.BottomBorder, actual.BottomBorderColor) })
                {
                    var expected = cell.GetProperty("Borders").GetProperty(edge); Assert.Equal(expected.GetProperty("Style").GetInt32(), kind);
                    if (kind != 0) Assert.Equal(expected.GetProperty("Color").GetString(), color ?? "FF000000");
                }
            }
        }
    }

    [Fact]
    public void HeaderOverlapNameAndStructuralGuardsAreAtomic()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new CreateTableCommand(id, new(0, 0, 2, 1), "Orders")); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromText("Rename")))])));
        Assert.Throws<ArgumentException>(() => session.Execute(new CreateTableCommand(id, new(0, 0, 2, 1), "Overlap")));
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 2, 1)));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SortRangeCommand(id, new(1, 1), new(2, 1), 1)));
        Assert.Throws<ArgumentException>(() => session.Execute(new CopySheetCommand(id, "Copy", 1)));
        Assert.Same(before, session.Snapshot);
        session.Execute(new SetCellsCommand([new(new(id, 1, 0), CellContent.FromLiteral(CellValue.FromText("body edit")))]));
        session.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(Bold: true))); Assert.True(session.CanUndo);
        using var limited = new DocumentSession(Fixture(), 2000);
        Assert.Throws<InvalidOperationException>(() => limited.Execute(new CreateTableCommand(limited.Snapshot.Sheets[0].Id, new(0, 0, 2, 1), "Orders")));
        Assert.False(limited.CanUndo); Assert.Empty(limited.Snapshot.Sheets[0].Data.Tables);
    }

    [Fact]
    public void MultipleTableFiltersProjectAcrossContextsAndKeepManualRows()
    {
        var source = Fixture(); var sheet = source.Sheets[0];
        var cells = sheet.Cells.SetItem(new(5, 0), T("Key")).SetItem(new(6, 0), T("yes")).SetItem(new(7, 0), T("no"));
        using var session = new DocumentSession(Book(new WorksheetSnapshot(sheet.Id, sheet.Name, cells, rows: [new(6, new(Hidden: true))])));
        session.Execute(new CreateTableCommand(sheet.Id, new(0, 0, 2, 1), "One")); session.Execute(new CreateTableCommand(sheet.Id, new(5, 0, 7, 0), "Two"));
        session.Execute(new SetTableFilterCommand(sheet.Id, 1, new(new(0, 0, 2, 1), [new(0, new(values: ["A"]))])));
        session.Execute(new SetTableFilterCommand(sheet.Id, 2, new(new(5, 0, 7, 0), [new(0, new(values: ["yes"]))])));
        var view = new SheetViewState(); view.Synchronize(session.Snapshot, sheet.Id, new(6, 0));
        Assert.Equal(2, view.TableId); Assert.Equal(2, view.HeaderRegions().Count()); Assert.True(view.IsFilteredOut(session.Snapshot, 0, 1)); Assert.True(view.IsFilteredOut(session.Snapshot, 0, 7));
        view.Synchronize(session.Snapshot, sheet.Id, new(15, 0)); Assert.Null(view.Region); Assert.True(view.IsFilteredOut(session.Snapshot, 0, 1));
        var path = FilePath("multiple.xlsx"); _storage.Save(path, session.Snapshot); var read = _storage.Open(path);
        Assert.Empty(read.FileState!.CompatibilityIssues); Assert.Equal(2, read.Sheets[0].Data.Tables.Length); Assert.True(read.Sheets[0].Rows[6].Hidden);
        Assert.Equal(new[] { 1, 7 }, WorksheetFilterEngine.Evaluate(read, read.Sheets[0].Id).Order());
        using var zip = ZipFile.OpenRead(path); Assert.NotNull(zip.GetEntry("xl/tables/table1.xml")); Assert.NotNull(zip.GetEntry("xl/tables/table2.xml"));
    }

    [Fact]
    public void NativeExcelTableGeometryAndDisplayMatchFixedIndependentEvidence()
    {
        using var oracle = ScalarFormulaFixture.Load("excel-table-basics"); var path = FilePath("native.xlsx");
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            foreach (var part in oracle.RootElement.GetProperty("Parts").EnumerateObject())
            { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open(), new UTF8Encoding(false)); writer.Write(part.Value.GetString()); }
        var snapshot = _storage.Open(path); Assert.False(snapshot.FileState!.FormulaReadOnly);
        foreach (var record in oracle.RootElement.GetProperty("After").EnumerateArray())
        {
            var sheet = snapshot.Sheets.Single(s => s.Name == record.GetProperty("Sheet").GetString()); var table = sheet.Data.Tables.Single();
            Assert.Equal(record.GetProperty("Name").GetString(), table.Name); Assert.Equal(record.GetProperty("Range").GetString(), table.Range.ToString());
            foreach (var cell in record.GetProperty("Cells").EnumerateArray())
            {
                var address = CellRectangle.Parse(cell.GetProperty("Address").GetString()!); var style = TablePresentation.GetStyle(snapshot, sheet, address.Top, address.Left);
                Assert.Equal(cell.GetProperty("Bold").GetBoolean(), style.Bold); Assert.Equal(cell.GetProperty("Font").GetString(), style.FontColor ?? "FF000000");
                Assert.Equal(cell.GetProperty("Fill").GetString(), style.FillColor ?? "FFFFFFFF");
                Assert.Equal(cell.GetProperty("Text").GetString(), CellDisplayFormatter.Format(sheet.GetCell(address.Top, address.Left)?.Value ?? CellValue.Blank, style, snapshot.DateSystem));
            }
        }
    }

    [Fact]
    public void ColumnDifferentialNumberFormatDrivesDisplayFilterAndOrdinarySave()
    {
        var source = Fixture(); var sheet = source.Sheets[0];
        var table = new WorksheetTable(1, "Orders", new(0, 0, 2, 1), [new(1, "Group"), new(2, "Amount")
        { DataStyle = new("wholeTable", NumberFormatId: 0, NumberFormatCode: "0.000") }], filter: new(new(0, 0, 2, 1), [new(1, new(values: ["10.000"]))]));
        var snapshot = Book(sheet.WithData(new() { Tables = [table] }));
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = FilePath("column-format." + extension); _storage.Save(path, snapshot); var read = _storage.Open(path); var actual = read.Sheets[0];
            Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(actual.Data.ContentEquals(snapshot.Sheets[0].Data));
            Assert.Equal("10.000", CellDisplayFormatter.Format(actual.GetCell(2, 1)!.Value, TablePresentation.GetStyle(read, actual, 2, 1), read.DateSystem));
            Assert.Equal(new[] { 1 }, WorksheetFilterEngine.Evaluate(read, actual.Id).Order());
            Assert.Equal(new[] { "10.000", "20.000" }, PersistentFiltering.Candidates(read, actual, 1, 1).Select(c => c.Label));
            if (extension == "xlsx")
            {
                using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry("xl/styles.xml")!.Open(); var xml = XElement.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                Assert.True((int)xml.Element(ns + "dxfs")!.Descendants(ns + "numFmt").Single().Attribute("numFmtId")! >= 164);
            }
        }
    }

    [Fact]
    public void InvalidTableDefinitionsAreRejectedBeforePublication()
    {
        Assert.Throws<ArgumentException>(() => new WorksheetTable(1, "DataTable", new(0, 0, 2, 1), [new(1, "A"), new(2, "a")]));
        Assert.Throws<ArgumentException>(() => new WorksheetTable(1, "DataTable", new(0, 0, 2, 0), [new(1, "A")], style: new() { Elements = [new("headerRow", FillColor: "red")] }));
        var first = new WorksheetTable(1, "First", new(0, 0, 2, 0), [new(1, "A")], style: new("Custom") { Elements = [new("headerRow", Bold: true)] });
        var second = new WorksheetTable(2, "Second", new(4, 0, 6, 0), [new(1, "A")], style: new("Custom") { Elements = [new("headerRow", Bold: false)] });
        Assert.Throws<ArgumentException>(() => Book(Fixture().Sheets[0].WithData(new() { Tables = [first, second] })));
    }

    [Fact]
    public void UnsupportedObjectsAndCalculatedColumnsRetainOriginalPackageProtection()
    {
        using var session = new DocumentSession(Fixture()); var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new CreateTableCommand(id, new(0, 0, 2, 1), "Orders"));
        var path = FilePath("unsupported.xlsx"); _storage.Save(path, session.Snapshot);
        void Edit(Action<XElement> action)
        {
            using var zip = ZipFile.Open(path, ZipArchiveMode.Update); var entry = zip.GetEntry("xl/tables/table1.xml")!;
            XElement xml; using (var stream = entry.Open()) xml = XElement.Load(stream); action(xml); entry.Delete(); using var output = zip.CreateEntry("xl/tables/table1.xml").Open(); xml.Save(output);
        }
        Edit(root => root.SetAttributeValue("connectionId", 1)); var bytes = File.ReadAllBytes(path); var read = _storage.Open(path);
        Assert.Empty(read.Sheets[0].Data.Tables); Assert.Contains(read.FileState!.CompatibilityDiagnostics, d => d.Feature == "table");
        _storage.Save(FilePath("original.xlsx"), read); Assert.Equal(bytes, File.ReadAllBytes(FilePath("original.xlsx")));
        using var modified = new DocumentSession(read); modified.Execute(new SetCellsCommand([new(new(read.Sheets[0].Id, 10, 0), CellContent.FromLiteral(CellValue.FromNumber(3)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(path, modified.Snapshot)); Assert.Equal(bytes, File.ReadAllBytes(path));
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Edit(root => { root.Attribute("connectionId")!.Remove(); root.Element(ns + "tableColumns")!.Elements().Last().Add(new XElement(ns + "calculatedColumnFormula", "Orders[@Amount]*2")); });
        var formula = _storage.Open(path); Assert.True(formula.FileState!.FormulaReadOnly); Assert.True(formula.Sheets[0].Data.Tables[0].HasFormulas);
        using var locked = new DocumentSession(formula);
        Assert.Throws<InvalidOperationException>(() => locked.Execute(new SetCellsCommand([new(new(formula.Sheets[0].Id, 1, 0), CellContent.FromLiteral(CellValue.FromText("blocked")))])));
    }
}
