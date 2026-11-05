using System.IO.Compression;
using System.Xml.Linq;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ColorFilterTests
{
    [Fact]
    public void IndependentExcelFilesRetainColorSemanticsAndUnsupportedPatternsRemainPreserved()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-color-filters.json")));
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "excel-color-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        foreach (var sample in fixture.RootElement.GetProperty("cases").EnumerateArray())
        {
            var path = Path.Combine(directory, sample.GetProperty("name").GetString() + ".xlsx");
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in sample.GetProperty("parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open()); writer.Write(part.Value.GetString()); }
            var storage = new TypedWorkbookStorage(); var book = storage.Open(path);
            var hidden = WorksheetFilterEngine.Evaluate(book, book.ActiveSheetId);
            Assert.Equal(sample.GetProperty("visibleRows").EnumerateArray().Select(x => x.GetInt32()), Enumerable.Range(2, 4).Where(row => !hidden.Contains(row - 1)));
            Assert.NotNull(book.Sheets[0].Data.Filter!.Columns[0].Color);
            Assert.DoesNotContain(book.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("autoFilter"));
            if (sample.GetProperty("name").GetString() != "fill") continue;
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/styles.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                xml.Root!.Element(ns + "dxfs")!.Descendants(ns + "patternFill").Single().SetAttributeValue("patternType", "darkGrid");
                entry.Delete(); using var output = zip.CreateEntry("xl/styles.xml").Open(); xml.Save(output);
            }
            var original = File.ReadAllBytes(path); var unsupported = storage.Open(path);
            Assert.Null(unsupported.Sheets[0].Data.Filter);
            Assert.Contains(unsupported.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("autoFilter"));
            var copy = Path.Combine(directory, "preserved.xlsx"); storage.Save(copy, unsupported); Assert.Equal(original, File.ReadAllBytes(copy));
        }
    }
    private static WorkbookSnapshot Fixture(bool font = false, string? color = "FFFF0000")
    {
        var source = Sheet("Colors", (0, 0, T("Amount")), (1, 0, new(CellContent.FromLiteral(CellValue.FromNumber(10), 1))),
            (2, 0, new(CellContent.FromLiteral(CellValue.FromNumber(20), 2))), (3, 0, new(CellContent.FromLiteral(CellValue.FromNumber(30), 3))), (4, 0, N(40)));
        var sheet = source.WithData(new(new(new(0, 0, 4, 0), [new(0, new(color: new(font, color)))])));
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { FillColor = "FFFF0000" }, new() { FillColor = "FF0000FF", FontColor = "FFFF0000" }, new() { FillColor = "FFFF0000", FontColor = "FFFF0000" }]);
    }
    [Theory]
    [InlineData(false, "FFFF0000", "2,4")]
    [InlineData(true, "FFFF0000", "1,4")]
    [InlineData(false, null, "1,2,3")]
    public void ColorFiltersPersistAndReapplyInBothFormats(bool font, string? color, string excluded)
    {
        var book = Fixture(font, color); var hidden = excluded.Split(',').Select(int.Parse).ToArray();
        Assert.Equal(hidden, WorksheetFilterEngine.Evaluate(book, book.ActiveSheetId).Order());
        var directory = Path.Combine(AppContext.BaseDirectory, "artifacts", "color-filter-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        foreach (var ext in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(directory, "color." + ext); var storage = new TypedWorkbookStorage(); storage.Save(path, book); var reopened = storage.Open(path);
            Assert.Equal(new AutoFilterColor(font, color), reopened.Sheets[0].Data.Filter!.Columns[0].Color);
            Assert.Equal(hidden, WorksheetFilterEngine.Evaluate(reopened, reopened.ActiveSheetId).Order());
            Assert.DoesNotContain(reopened.FileState!.CompatibilityDiagnostics, d => d.Feature.Contains("autoFilter"));
        }
    }
    [Fact]
    public void ColorChangesRefreshSubtotalAndUndoRestoresTheFilterResult()
    {
        var original = Fixture(); var sheet = original.Sheets[0]; var book = original.WithSheets([sheet.WithCells(sheet.Cells.Add(new(6, 0), F("=SUBTOTAL(9,A2:A5)")))]);
        using var session = new DocumentSession(book);
        double Calculate()
        {
            var request = session.BeginCalculation(); var result = new WorkbookCalculator(new NoScripts()).Calculate(request); Assert.True(session.TryCommitCalculation(result));
            return session.Snapshot.Sheets[0].GetCell(6, 0)!.Value.NumberValue;
        }
        Assert.Equal(40, Calculate());
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 2, 0)], new(FillColor: "FFFF0000"))); Assert.Equal(60, Calculate());
        Assert.True(session.Undo()); Assert.Equal(40, Calculate());
    }
    [Fact]
    public void MultipleColumnsCombineAndManualHiddenStateSurvives()
    {
        var book = Fixture(); var source = book.Sheets[0]; var sheet = new WorksheetSnapshot(source.Id, source.Name, source.Cells, rows: [new(3, new(Hidden: true))], data: source.Data);
        var filtered = book.WithSheets([sheet]); Assert.Contains(2, WorksheetFilterEngine.Evaluate(filtered, sheet.Id)); Assert.True(sheet.Rows[3].Hidden);
        using var session = new DocumentSession(filtered);
        session.Execute(new SetAutoFilterCommand(sheet.Id, new(new(0, 0, 4, 1), [new(0, new(color: new(false, "FFFF0000"))), new(1, new(values: ["x"]))])));
        Assert.Equal(4, WorksheetFilterEngine.Evaluate(session.Snapshot, sheet.Id).Count);
        session.Undo(); Assert.Equal(2, WorksheetFilterEngine.Evaluate(session.Snapshot, sheet.Id).Count);
    }
    [Fact]
    public void InvalidColorAndMixedConditionsCannotEnterTheModel()
    {
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(color: new(false, "00FF0000")));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(values: ["x"], color: new(false, "FFFF0000")));
        Assert.Throws<ArgumentException>(() => new AutoFilterColumn(color: new(true, null)));
        Assert.Equal("FFFF0000", new AutoFilterColumn(color: new(false, "ffff0000")).WithButton(false).Color!.Argb);
    }
    [Fact]
    public void ConditionalDisplayRulesRejectColorFilteringInsteadOfReinterpretingDisplayedColors()
    {
        var book = Fixture(); var sheet = book.Sheets[0];
        var rule = new ConditionalFormatRule([new(1, 0, 4, 0)], ConditionalFormatKind.Expression, 1, new(FillColor: "FF00FF00"), "=TRUE");
        var next = book.WithSheets([sheet.WithData(sheet.Data with { ConditionalFormats = [rule] })]);
        Assert.Throws<InvalidOperationException>(() => WorksheetFilterEngine.Evaluate(next, sheet.Id));
    }
    [Fact]
    public void ProduceColorFilesForIndependentExcelReapplication()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "artifacts", "color-filter-oracle"); Directory.CreateDirectory(path);
        var storage = new TypedWorkbookStorage();
        storage.Save(Path.Combine(path, "fill.xlsx"), Fixture()); storage.Save(Path.Combine(path, "font.xlsx"), Fixture(true)); storage.Save(Path.Combine(path, "none.xlsx"), Fixture(color: null));
    }
}
