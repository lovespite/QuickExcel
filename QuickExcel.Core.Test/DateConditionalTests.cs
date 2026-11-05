using System.Collections.Immutable;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DateConditionalTests
{
    [Fact]
    public void IndependentExcelAllTenPeriodsAndBothEpochsMatchAndPersist()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-date-conditional.json")));
        var today = DateTime.Parse(fixture.RootElement.GetProperty("today").GetString()!, System.Globalization.CultureInfo.InvariantCulture);
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "date-conditional-oracle"); Directory.CreateDirectory(directory);
        var imported = Path.Combine(directory, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(imported);
        foreach (var sample in fixture.RootElement.GetProperty("books").EnumerateArray())
        {
            var systemNumber = sample.GetProperty("system").GetInt32(); var system = systemNumber == 1900 ? ExcelDateSystem.Windows1900 : ExcelDateSystem.Mac1904;
            var path = Path.Combine(imported, systemNumber + ".xlsx");
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in sample.GetProperty("parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open()); writer.Write(part.Value.GetString()); }
            foreach (var spec in fixture.RootElement.GetProperty("cases").EnumerateArray().Where(c => c.GetProperty("system").GetInt32() == systemNumber))
            {
                var period = spec.GetProperty("period").GetInt32();
                Assert.Equal(FormulaSyntax.Print(FormulaSyntax.Parse(spec.GetProperty("formula").GetString()!)), FormulaSyntax.Print(FormulaSyntax.Parse(ConditionalDates.Formula((ConditionalDatePeriod)period, new(0, period)))));
            }
            var storage = new TypedWorkbookStorage(); var excel = storage.Open(path);
            Assert.True(excel.Sheets[0].Data.ConditionalFormats.Length == 10, string.Join("\n", excel.FileState!.CompatibilityDiagnostics.Where(d => d.Feature.Contains("conditionalFormatting"))));
            Assert.DoesNotContain(excel.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("conditionalFormatting"));
            var cells = new Dictionary<CellPosition, CellSnapshot>(); var serial = ExcelSerialDate.FromDateTime(today, system).Serial;
            for (var column = 0; column < 10; column++)
            {
                for (var row = 0; row < 91; row++) cells[new(row, column)] = N(serial + row - 45 + 0.5);
                cells[new(91, column)] = T(serial.ToString(System.Globalization.CultureInfo.InvariantCulture)); cells[new(92, column)] = T(today.ToString("yyyy-MM-dd"));
                cells[new(94, column)] = new(CellContent.FromLiteral(CellValue.FromBoolean(true))); cells[new(95, column)] = F("=1/0");
            }
            var rules = Enumerable.Range(0, 10).Select(i => new ConditionalFormatRule([new(0, i, 95, i)], ConditionalFormatKind.TimePeriod, i + 1, new(FillColor: "FFFF0000"), datePeriod: (ConditionalDatePeriod)i));
            var sheet = new WorksheetSnapshot(SheetId.New(), "Dates", cells, data: new() { ConditionalFormats = rules.ToImmutableArray() });
            using var session = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, system, [sheet]));
            Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
            void Check(WorkbookSnapshot book)
            {
                foreach (var spec in fixture.RootElement.GetProperty("cases").EnumerateArray().Where(c => c.GetProperty("system").GetInt32() == systemNumber))
                {
                    var column = spec.GetProperty("period").GetInt32();
                    var actual = Enumerable.Range(0, 96).Where(row => ConditionalPresentation.GetStyleAtDate(book, book.Sheets[0], row, column, today).FillColor == "FFFF0000").Select(r => r + 1);
                    Assert.Equal(spec.GetProperty("matches").EnumerateArray().Select(v => v.GetInt32()), actual);
                }
            }
            Check(excel); Check(session.Snapshot);
            foreach (var ext in new[] { "xlsx", "xlsq" })
            { var output = Path.Combine(directory, systemNumber + "." + ext); storage.Save(output, session.Snapshot); Check(storage.Open(output)); }
        }
    }
    [Fact]
    public void CalendarDayInvalidatesUnchangedSnapshotAndBoundariesIncludeYearTransitions()
    {
        var date = new DateTime(2026, 12, 31); var serial = ExcelSerialDate.FromDateTime(date, ExcelDateSystem.Windows1900).Serial;
        var sheet = Sheet("S", (0, 0, N(serial + .75))).WithData(new() { ConditionalFormats = [new([new(0, 0, 0, 0)], ConditionalFormatKind.TimePeriod, 1, new(Bold: true), datePeriod: ConditionalDatePeriod.Today)] });
        var book = Book(sheet);
        Assert.True(ConditionalPresentation.GetStyleAtDate(book, sheet, 0, 0, date).Bold);
        Assert.False(ConditionalPresentation.GetStyleAtDate(book, sheet, 0, 0, date.AddDays(1)).Bold);
        Assert.True(ConditionalDates.Matches(serial, book.DateSystem, ConditionalDatePeriod.LastMonth, new(2027, 1, 1)));
        Assert.True(ConditionalDates.Matches(serial, book.DateSystem, ConditionalDatePeriod.NextMonth, new(2026, 11, 30)));
        var sunday = new DateTime(2027, 1, 3); var sundaySerial = ExcelSerialDate.FromDateTime(sunday, book.DateSystem).Serial;
        Assert.True(ConditionalDates.Matches(sundaySerial, book.DateSystem, ConditionalDatePeriod.ThisWeek, sunday.AddDays(6)));
        Assert.False(ConditionalDates.Matches(sundaySerial - 1, book.DateSystem, ConditionalDatePeriod.ThisWeek, sunday));
    }
    [Fact]
    public void RuleIdentityPriorityUndoAndStructureEditsRetainTheDatePeriod()
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        var rule = new ConditionalFormatRule([new(1, 0, 4, 0)], ConditionalFormatKind.TimePeriod, 1, new(Bold: true), stopIfTrue: true, datePeriod: ConditionalDatePeriod.NextMonth);
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [rule])); Assert.True(session.Undo()); Assert.Empty(session.Snapshot.Sheets[0].Data.ConditionalFormats);
        Assert.True(session.Redo()); session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1, false));
        var changed = session.Snapshot.Sheets[0].Data.ConditionalFormats.Single();
        Assert.Equal(rule.ScriptId, changed.ScriptId); Assert.Equal(rule.DatePeriod, changed.DatePeriod); Assert.True(changed.StopIfTrue);
        Assert.Equal(new(2, 0, 5, 0), changed.Ranges.Single());
        Assert.Throws<ArgumentException>(() => new ConditionalFormatRule(rule.Ranges, rule.Kind, 1, rule.Style, "=TRUE", datePeriod: ConditionalDatePeriod.Today));
    }
    [Fact]
    public void ModifiedTimePeriodFormulaIsPreservedRatherThanSilentlyReinterpreted()
    {
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "date-cf-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "source.xlsx"); var storage = new TypedWorkbookStorage();
        storage.Save(path, Book(Sheet("S").WithData(new() { ConditionalFormats = [new([new(0, 0, 0, 0)], ConditionalFormatKind.TimePeriod, 1, new(Bold: true), datePeriod: ConditionalDatePeriod.Today)] })));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument xml; using (var input = entry.Open()) xml = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"; xml.Descendants(ns + "formula").Single().Value = "TRUE";
            entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
        }
        var bytes = File.ReadAllBytes(path); var book = storage.Open(path); Assert.Empty(book.Sheets[0].Data.ConditionalFormats);
        Assert.Contains(book.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("conditionalFormatting"));
        var copy = Path.Combine(folder, "copy.xlsx"); storage.Save(copy, book); Assert.Equal(bytes, File.ReadAllBytes(copy));
    }
}
