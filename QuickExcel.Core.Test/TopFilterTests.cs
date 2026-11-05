using System.IO.Compression;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class TopFilterTests
{
    private static WorkbookSnapshot Fixture(AutoFilterTop top, bool group = false, bool table = false)
    {
        var sheet = Sheet("Ranks", (0, 0, T("Amount")), (0, 1, T("Group")), (1, 0, N(10)), (2, 0, N(20)), (3, 0, N(20)), (4, 0, N(30)), (5, 0, N(40)),
            (6, 0, T("999")), (8, 0, T("tail")), (9, 0, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))));
        var cells = sheet.Cells.ToBuilder(); for (var row = 1; row <= 9; row++) cells[new(row, 1)] = T(row <= 3 ? "A" : "B");
        var columns = new Dictionary<int, AutoFilterColumn> { [0] = new(top: top) }; if (group) columns.Add(1, new(values: ["A"]));
        var filter = new WorksheetAutoFilter(new(0, 0, 9, 1), columns);
        var data = table ? new WorksheetData { Tables = [new(1, "RanksTable", filter.Range, [new(1, "Amount"), new(2, "Group")], filter: filter)] } : new(filter);
        return Book(sheet.WithCells(cells.ToImmutable()).WithData(data));
    }
    [Fact]
    public void IndependentExcelCasesMatchAndProduceFilesForReapplication()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-top-filters.json")));
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "top-filter-oracle"); Directory.CreateDirectory(directory);
        var imported = Path.Combine(directory, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(imported);
        foreach (var sample in fixture.RootElement.GetProperty("cases").EnumerateArray())
        {
            var spec = sample.GetProperty("spec"); var name = spec.GetProperty("name").GetString()!;
            var top = new AutoFilterTop(spec.GetProperty("count").GetInt32(), spec.GetProperty("percent").GetBoolean(), spec.GetProperty("bottom").GetBoolean());
            var expected = sample.GetProperty("visibleRows").EnumerateArray().Select(v => v.GetInt32()).ToArray();
            var book = Fixture(top, spec.GetProperty("group").GetBoolean(), spec.GetProperty("table").GetBoolean());
            void Check(WorkbookSnapshot value)
            {
                var hidden = WorksheetFilterEngine.Evaluate(value, value.ActiveSheetId);
                Assert.Equal(expected, Enumerable.Range(2, 9).Where(row => !hidden.Contains(row - 1)));
                Assert.Equal(top, value.Sheets[0].Data.AllFilters.Single().Columns[0].Top);
            }
            Check(book); var storage = new TypedWorkbookStorage();
            foreach (var ext in new[] { "xlsx", "xlsq" })
            { var file = Path.Combine(directory, name + "." + ext); storage.Save(file, book); Check(storage.Open(file)); }
            var source = Path.Combine(imported, name + ".xlsx");
            using (var zip = ZipFile.Open(source, ZipArchiveMode.Create))
                foreach (var part in sample.GetProperty("parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open()); writer.Write(part.Value.GetString()); }
            var excel = storage.Open(source); Check(excel);
            Assert.DoesNotContain(excel.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("autoFilter"));
        }
    }
    [Fact]
    public void SubtotalTracksThresholdInputsOutsideItsOwnValueRange()
    {
        var book = Fixture(new(2)); var sheet = book.Sheets[0];
        book = book.WithSheets([sheet.WithCells(sheet.Cells.Add(new(12, 0), F("=SUBTOTAL(9,A2:A3)")))]);
        using var session = new DocumentSession(book); var calculator = new WorkbookCalculator(new NoScripts());
        double Calculate()
        { Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation()))); return session.Snapshot.Sheets[0].GetCell(12, 0)!.Value.NumberValue; }
        Assert.Equal(0, Calculate());
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 5, 0), CellContent.FromLiteral(CellValue.FromNumber(0)))]));
        Assert.Equal(20, Calculate()); Assert.True(session.Undo()); Assert.Equal(0, Calculate());
    }
    [Fact]
    public void ErrorsOldCachesInvalidModesAndCancellationRejectAtomically()
    {
        var book = Fixture(new(2)); var sheet = book.Sheets[0];
        var invalid = book.WithSheets([sheet.WithCells(sheet.Cells.SetItem(new(7, 0), new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.DivisionByZero))))) ]);
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.Evaluate(invalid, sheet.Id));
        var stale = book.WithSheets([sheet.WithCells(sheet.Cells.SetItem(new(7, 0), F("=999")))]);
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.Evaluate(stale, sheet.Id));
        Assert.Throws<OperationCanceledException>(() => WorksheetFilterEngine.Evaluate(book, sheet.Id, cancellationToken: new(true)));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(top: new(101, true)));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(values: ["x"], top: new(10)));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(color: new(false, "FFFF0000"), top: new(10)));
    }
    [Fact]
    public void ManualHiddenRowsRemainIndependentAndNonNumericColumnsAreRejected()
    {
        var book = Fixture(new(2)); var source = book.Sheets[0];
        var sheet = new WorksheetSnapshot(source.Id, source.Name, source.Cells, rows: [new(5, new(Hidden: true))], data: source.Data);
        var hidden = WorksheetFilterEngine.Evaluate(book.WithSheets([sheet]), sheet.Id);
        Assert.DoesNotContain(5, hidden); Assert.True(sheet.Rows[5].Hidden); Assert.DoesNotContain(4, hidden);
        var text = Sheet("Text", (0, 0, T("Header")), (1, 0, T("999"))).WithData(new(new(new(0, 0, 1, 0), [new(0, new(top: new(1)))])));
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.Evaluate(Book(text), text.Id));
    }
}
