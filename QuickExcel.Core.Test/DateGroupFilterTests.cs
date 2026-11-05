using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DateGroupFilterTests
{
    private static WorkbookSnapshot Fixture(ExcelDateSystem system, string kind, bool table = false)
    {
        var dates = new[] { new DateTime(2026, 11, 5, 12, 30, 0), new(2026, 11, 30), new(2026, 12, 1), new(2025, 11, 5), new(2026, 11, 5, 12, 30, 0) };
        var source = Sheet("Dates", (0, 0, T("Date")), (6, 0, T("2026-11-05")), (8, 0, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            (9, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.DivisionByZero)))));
        var cells = source.Cells.ToBuilder();
        for (var i = 0; i < dates.Length; i++) cells[new(i + 1, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(ExcelSerialDate.FromDateTime(dates[i], system).Serial), i < 4 ? 1 : 0));
        AutoFilterDateGroup[] groups = kind switch { "year" => [new(2026)], "month" => [new(2026, 11)], "day" => [new(2026, 11, 5)], _ => [new(2026, 12, 1), new(2025)] };
        var filter = new WorksheetAutoFilter(new(0, 0, 9, 0), [new(0, new(dateGroups: groups))]);
        var data = table ? new WorksheetData { Tables = [new(1, "DateTable", filter.Range, [new(1, "Date")], filter: filter)] } : new(filter);
        return new(Guid.NewGuid(), 0, system, [source.WithCells(cells.ToImmutable()).WithData(data)], styles: [new(), new() { NumberFormatCode = "yyyy-mm-dd" }]);
    }

    [Fact]
    public void IndependentExcelYearMonthDayGroupsMatchBothDateSystemsAndPersist()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-date-filters.json")));
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "date-filter-oracle"); Directory.CreateDirectory(directory);
        var imported = Path.Combine(directory, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(imported);
        foreach (var sample in fixture.RootElement.GetProperty("cases").EnumerateArray())
        {
            var name = sample.GetProperty("name").GetString()!;
            var expected = sample.GetProperty("visibleRows").EnumerateArray().Select(v => v.GetInt32()).ToArray();
            var system = sample.GetProperty("system").GetInt32() == 1900 ? ExcelDateSystem.Windows1900 : ExcelDateSystem.Mac1904;
            void Check(WorkbookSnapshot value)
            {
                var hidden = WorksheetFilterEngine.Evaluate(value, value.ActiveSheetId, allowSourceCaches: true);
                Assert.Equal(expected, Enumerable.Range(2, 9).Where(row => !hidden.Contains(row - 1)));
                Assert.NotEmpty(value.Sheets[0].Data.AllFilters.Single().Columns[0].DateGroups);
                if (value.FileState is { } state) Assert.DoesNotContain(state.CompatibilityDiagnostics, d => d.Feature.Contains("autoFilter"));
            }
            var storage = new TypedWorkbookStorage();
            foreach (var table in new[] { false, true })
            {
                var book = Fixture(system, sample.GetProperty("kind").GetString()!, table); Check(book);
                foreach (var ext in new[] { "xlsx", "xlsq" })
                { var file = Path.Combine(directory, name + (table ? "-table" : "") + "." + ext); storage.Save(file, book); Check(storage.Open(file)); }
            }
            var source = Path.Combine(imported, name + ".xlsx");
            using (var zip = ZipFile.Open(source, ZipArchiveMode.Create))
                foreach (var part in sample.GetProperty("parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open()); writer.Write(part.Value.GetString()); }
            Check(storage.Open(source));
        }
    }

    [Fact]
    public void MixedLabelsBlanksLeapDayAndSubDayValuesHaveExplicitSemantics()
    {
        var rule = new AutoFilterColumn(values: ["keep"], includeBlank: true, dateGroups: [new(1900, 2, 29), new(2026, 11)]);
        var match = WorksheetFilterEngine.Compile(rule); var style = new StyleDefinition { NumberFormatCode = "yyyy-mm-dd" };
        Assert.True(match(CellValue.FromNumber(60.999999), style, ExcelDateSystem.Windows1900));
        Assert.False(match(CellValue.FromNumber(60), style, ExcelDateSystem.Mac1904));
        Assert.False(match(CellValue.FromNumber(60), new(), ExcelDateSystem.Windows1900));
        Assert.False(match(CellValue.FromNumber(-1), style, ExcelDateSystem.Windows1900));
        Assert.False(match(CellValue.FromText("1900-02-29"), style, ExcelDateSystem.Windows1900));
        Assert.True(match(CellValue.FromText("keep"), style, ExcelDateSystem.Windows1900));
        Assert.True(match(CellValue.Blank, style, ExcelDateSystem.Windows1900));
        Assert.True(rule.ContentEquals(rule.WithButton(true))); Assert.True(rule.DateGroups.SequenceEqual(rule.WithButton(false).DateGroups));
        Assert.Throws<ArgumentException>(() => AutoFilterDateGroup.Parse("2026-02-29"));
        Assert.Throws<ArgumentException>(() => AutoFilterDateGroup.Parse("26-11"));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(dateGroups: [new(2026)], top: new(5)));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(dateGroups: Enumerable.Range(0, 10_001).Select(i => new AutoFilterDateGroup(2000 + i / 372, 1 + i / 31 % 12, 1 + i % 28))));
    }

    [Fact]
    public void ReapplyUndoAndStaleResultsAreAtomicAndManualHidingIsIndependent()
    {
        var book = Fixture(ExcelDateSystem.Windows1900, "month"); var sheet = book.Sheets[0];
        using var session = new DocumentSession(book);
        session.Execute(new SetAutoFilterCommand(sheet.Id, new(sheet.Data.Filter!.Range, [new(0, new(dateGroups: [new(2025)]))])));
        Assert.DoesNotContain(4, WorksheetFilterEngine.Evaluate(session.Snapshot, sheet.Id));
        Assert.True(session.Undo()); Assert.Contains(4, WorksheetFilterEngine.Evaluate(session.Snapshot, sheet.Id));
        Assert.True(session.Redo());
        var stale = book.WithSheets([sheet.WithCells(sheet.Cells.SetItem(new(1, 0), F("=1")))]);
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.Evaluate(stale, sheet.Id));
        Assert.Throws<OperationCanceledException>(() => WorksheetFilterEngine.Evaluate(book, sheet.Id, cancellationToken: new(true)));
        var manual = new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells, rows: [new(1, new(Hidden: true))], data: sheet.Data);
        Assert.DoesNotContain(1, WorksheetFilterEngine.Evaluate(book.WithSheets([manual]), sheet.Id));
        Assert.True(manual.Rows[1].Hidden);
    }

    [Fact]
    public void UnsupportedDateGranularityRetainsOriginalPackage()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "date-filter-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "hour.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, Fixture(ExcelDateSystem.Windows1900, "day"));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var group = xml.Descendants(ns + "dateGroupItem").Single(); group.SetAttributeValue("dateTimeGrouping", "hour"); group.SetAttributeValue("hour", 12);
            entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
        }
        var original = File.ReadAllBytes(path); var unsupported = storage.Open(path);
        Assert.Null(unsupported.Sheets[0].Data.Filter); Assert.Contains(unsupported.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("autoFilter"));
        var copy = Path.Combine(directory, "preserved.xlsx"); storage.Save(copy, unsupported); Assert.Equal(original, File.ReadAllBytes(copy));
    }
}
