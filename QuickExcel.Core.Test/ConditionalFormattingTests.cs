using System.IO.Compression;
using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ConditionalFormattingTests
{
    private static StyleDefinition Style(WorkbookSnapshot book, int row = 0, int column = 0, int sheet = 0) =>
        ConditionalPresentation.GetStyle(book, book.Sheets[sheet], row, column);
    private static readonly DifferentialStyle Red = new(FillColor: "FFFF0000");
    [Fact]
    public void PriorityCombinesOnlySpecifiedPropertiesAndStopSuppressesLowerRules()
    {
        var sheet = Sheet("Data", (0, 0, N(3)));
        var high = new ConditionalFormatRule([new(0, 0, 0, 0)], ConditionalFormatKind.CellIs, 1, Red with { Bold = false }, "=1", comparison: ConditionalComparison.GreaterThan);
        var low = new ConditionalFormatRule(high.Ranges, ConditionalFormatKind.Expression, 2, new(Bold: true, Italic: true, FontColor: "FF0000FF", Bottom: new(1, "FF000000")), "=TRUE");
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [high, low]));
        var style = Style(session.Snapshot);
        Assert.Equal("FFFF0000", style.FillColor); Assert.False(style.Bold); Assert.True(style.Italic); Assert.Equal("FF0000FF", style.FontColor);
        Assert.Equal(1, Style(session.Snapshot, 1).TopBorder); // shared edge outside the rule
        var stop = new ConditionalFormatRule(high.Ranges, high.Kind, 1, high.Style, high.Formula1, comparison: high.Comparison, stopIfTrue: true);
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [stop, low]));
        Assert.False(Style(session.Snapshot).Italic); Assert.Equal(0, Style(session.Snapshot, 1).TopBorder);
        session.Undo(); Assert.True(Style(session.Snapshot).Italic); session.Undo(); Assert.Empty(session.Snapshot.Sheets[0].Data.ConditionalFormats);
        session.Redo(); Assert.Equal("FFFF0000", Style(session.Snapshot).FillColor);
    }
    [Fact]
    public void ViewProjectionUsesCandidateDependenciesAndInvalidatesOnEveryImmutableSnapshot()
    {
        var sheet = Sheet("Data", (0, 0, N(1)), (0, 1, F("=A1*2")), (0, 2, N(5)));
        var rule = new ConditionalFormatRule([new(0, 2, 0, 2)], ConditionalFormatKind.Expression, 1, Red, "=$B$1>3");
        using var session = new DocumentSession(Book(sheet.WithData(new() { ConditionalFormats = [rule] })));
        Assert.Null(Style(session.Snapshot, 0, 2).FillColor);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
        Assert.Equal("FFFF0000", Style(session.Snapshot, 0, 2).FillColor);
        session.Undo(); Assert.Null(Style(session.Snapshot, 0, 2).FillColor);
        var invalid = new ConditionalFormatRule(rule.Ranges, ConditionalFormatKind.Expression, 1, Red, "=1/0");
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [invalid])); Assert.Null(Style(session.Snapshot, 0, 2).FillColor);
        Assert.Empty(ConditionalPresentation.Diagnostics(session.Snapshot));
    }
    [Fact]
    public void DeletingAnchorAndWholeFirstAreaTranslatesSurvivingRelativeFormula()
    {
        var sheet = Sheet("Data");
        var rule = new ConditionalFormatRule([new(1, 1, 4, 1), new(7, 1, 9, 1)], ConditionalFormatKind.Expression, 1, Red, "=$A2>2");
        using var session = new DocumentSession(Book(sheet.WithData(new() { ConditionalFormats = [rule] })));
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1, true));
        var changed = session.Snapshot.Sheets[0].Data.ConditionalFormats.Single();
        Assert.Equal("=($A2>2)", changed.Formula1); Assert.Equal(new(1, 1, 3, 1), changed.Ranges[0]);
        session.Undo(); session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 4, true));
        changed = session.Snapshot.Sheets[0].Data.ConditionalFormats.Single();
        Assert.Equal(new(3, 1, 5, 1), Assert.Single(changed.Ranges)); Assert.Equal("=($A4>2)", changed.Formula1);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Column, 1, 1, true));
        Assert.Empty(session.Snapshot.Sheets[0].Data.ConditionalFormats);
    }
    [Fact]
    public void NamesAndSheetCopyKeepIndependentRulesAndBoundReferences()
    {
        var sheet = Sheet("Data", (0, 0, N(2)));
        var rule = new ConditionalFormatRule([new(0, 1, 2, 1)], ConditionalFormatKind.Expression, 1, Red, "=Data!A1>Threshold");
        using var session = new DocumentSession(Book(sheet.WithData(new() { ConditionalFormats = [rule] })).WithDefinedNames([new("Threshold", "=1")]));
        session.Execute(new RenameDefinedNameCommand("Threshold", null, "Minimum"));
        session.Execute(new RenameSheetCommand(sheet.Id, "Input"));
        session.Execute(new CopySheetCommand(sheet.Id, "Copy", 1));
        Assert.Contains("'Input'!", session.Snapshot.Sheets[0].Data.ConditionalFormats[0].Formula1);
        Assert.Contains("'Copy'!", session.Snapshot.Sheets[1].Data.ConditionalFormats[0].Formula1);
        Assert.Contains("Minimum", session.Snapshot.Sheets[1].Data.ConditionalFormats[0].Formula1);
        session.Execute(new SetConditionalFormatsCommand(session.Snapshot.Sheets[1].Id, []));
        Assert.Single(session.Snapshot.Sheets[0].Data.ConditionalFormats);
    }
    [Fact]
    public void UnsupportedDependencyIsReportedWithoutRunningScriptAndResourceLimitIsVisible()
    {
        var sheet = Sheet("Data", (0, 0, F("=CUSTOM_FN(1)")));
        var rule = new ConditionalFormatRule([new(0, 1, 0, 1)], ConditionalFormatKind.Expression, 1, Red, "=A1>0");
        var book = Book(sheet.WithData(new() { ConditionalFormats = [rule] }));
        Assert.Null(Style(book, 0, 1).FillColor); Assert.NotEmpty(ConditionalPresentation.Diagnostics(book));
        var large = new ConditionalFormatRule([new(0, 0, WorkbookLimits.MaxRows - 1, 29)], ConditionalFormatKind.DuplicateValues, 1, Red);
        book = Book(Sheet("Large", (0, 0, N(1))).WithData(new() { ConditionalFormats = [large] }));
        Assert.Null(Style(book).FillColor); Assert.Contains(ConditionalPresentation.Diagnostics(book), s => s.Contains("2,000,000"));
    }
    [Fact]
    public void UnknownRulePreservesOriginalBytesAndCannotBeOverwrittenByRuleEditor()
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-cf-" + Guid.NewGuid().ToString("N") + ".xlsx"); var output = path + ".xlsx";
        try
        {
            var storage = new TypedWorkbookStorage();
            var sheet = Sheet("Data").WithData(new() { ConditionalFormats = [new([new(0, 0, 0, 0)], ConditionalFormatKind.Expression, 1, Red, "=TRUE")] });
            storage.Save(path, Book(sheet));
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; string xml;
                using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
                entry.Delete(); using var writer = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open(), new UTF8Encoding(false));
                writer.Write(xml.Replace("type=\"expression\"", "type=\"iconSet\""));
            }
            var source = File.ReadAllBytes(path); var imported = storage.Open(path);
            Assert.Empty(imported.Sheets[0].Data.ConditionalFormats);
            Assert.Contains(imported.FileState!.CompatibilityDiagnostics, d => d.Feature == "worksheet/conditionalFormatting");
            using var session = new DocumentSession(imported);
            Assert.Throws<NotSupportedException>(() => session.Execute(new SetConditionalFormatsCommand(imported.Sheets[0].Id, [])));
            storage.Save(output, imported); Assert.Equal(source, File.ReadAllBytes(output));
        }
        finally { File.Delete(path); File.Delete(output); }
    }
    [Fact]
    public void RuleCollectionLimitsAndReorderingAreAtomic()
    {
        var sheet = Sheet("Data"); using var session = new DocumentSession(Book(sheet));
        var one = new ConditionalFormatRule([new(0, 0, 0, 0)], ConditionalFormatKind.Expression, 1, Red, "=TRUE");
        Assert.Throws<ArgumentException>(() => session.Execute(new SetConditionalFormatsCommand(sheet.Id, [one, one])));
        Assert.Empty(session.Snapshot.Sheets[0].Data.ConditionalFormats);
        Assert.Throws<NotSupportedException>(() => new ConditionalFormatRule(one.Ranges, one.Kind, 1, Red, "=RAND()>0.5"));
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [one]));
        Assert.Throws<NotSupportedException>(() => session.Execute(new MoveAxisCommand(sheet.Id, SheetAxis.Row, 0, 1, 3)));
    }
    [Fact]
    public void ClipboardAndFillCannotSilentlyDiscardOrSplitConditionalRules()
    {
        var sheet = Sheet("Data", (0, 0, N(2)));
        var rule = new ConditionalFormatRule([new(0, 0, 2, 0)], ConditionalFormatKind.Expression, 1, Red, "=TRUE");
        var book = Book(sheet.WithData(new() { ConditionalFormats = [rule] })); using var session = new DocumentSession(book);
        var clipboard = WorkbookClipboard.Capture(book, sheet.Id, new(0, 0, 0, 0));
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(clipboard, sheet.Id, new(0, 1, 0, 1), book.Revision)));
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteCellsCommand(clipboard, sheet.Id, new(0, 1, 0, 1), book.Revision, CellPasteMode.Formats)));
        Assert.Throws<NotSupportedException>(() => session.Execute(new FillCellsCommand(sheet.Id, new(0, 0, 0, 0), new(0, 0, 2, 0))));
        Assert.Throws<NotSupportedException>(() => session.Execute(new ClearCellFormatsCommand([new(sheet.Id, 0, 0)])));
        Assert.Same(book, session.Snapshot);
        session.Execute(new PasteCellsCommand(clipboard, sheet.Id, new(0, 1, 0, 1), book.Revision, CellPasteMode.Values));
        Assert.Equal(CellValue.FromNumber(2), session.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
        Assert.True(rule.ContentEquals(session.Snapshot.Sheets[0].Data.ConditionalFormats[0]));
    }
    [Fact]
    public void WholeRecordSortRetainsColumnRulesAndRejectsFragmentingOnlySomeRuleRows()
    {
        var sheet = Sheet("Data", (0, 0, N(3)), (1, 0, N(1)), (2, 0, N(2)), (0, 1, N(30)), (1, 1, N(10)), (2, 1, N(20)));
        var rule = new ConditionalFormatRule([new(0, 1, 2, 1)], ConditionalFormatKind.Expression, 1, Red, "=$A1>1");
        using var session = new DocumentSession(Book(sheet.WithData(new() { ConditionalFormats = [rule] })));
        session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(2, 1), 0));
        Assert.Null(Style(session.Snapshot, 0, 1).FillColor); Assert.Equal("FFFF0000", Style(session.Snapshot, 2, 1).FillColor);
        Assert.True(rule.ContentEquals(session.Snapshot.Sheets[0].Data.ConditionalFormats[0]));
        Assert.Throws<NotSupportedException>(() => session.Execute(new SortRangeCommand(sheet.Id, new(0, 0), new(3, 1), 0)));
    }
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void EditableRuleMetadataSurvivesOrdinarySaveAndReopen(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-cf-" + Guid.NewGuid().ToString("N") + "." + extension);
        try
        {
            var rule = new ConditionalFormatRule([new(0, 0, 0, 0), new(2, 0, 3, 0)], ConditionalFormatKind.CellIs, 1,
                new(Bold: false, Italic: true, Strike: true, Underline: "single", Bottom: new(1, "FF0000FF"), NumberFormatId: 0, NumberFormatCode: "0.00%"),
                "=0", "=10", stopIfTrue: true);
            var book = Book(Sheet("Data", (0, 0, N(2))).WithData(new() { ConditionalFormats = [rule] }));
            var storage = new TypedWorkbookStorage(); storage.Save(path, book); var opened = storage.Open(path);
            Assert.True(rule.ContentEquals(Assert.Single(opened.Sheets[0].Data.ConditionalFormats)));
            Assert.Equal(Style(book), Style(opened));
        }
        finally { File.Delete(path); }
    }
    [Fact]
    public void NativeExcelFixtureMatchesDisplayedPropertiesIncludingWildcardDuplicates()
    {
        using var fixture = ScalarFormulaFixture.Load("excel-conditional-formatting"); var root = fixture.RootElement;
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-cf-native-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
                foreach (var part in root.GetProperty("Parts").EnumerateObject())
                { using var writer = new StreamWriter(zip.CreateEntry(part.Name).Open(), new UTF8Encoding(false)); writer.Write(part.Value.GetString()); }
            var book = new TypedWorkbookStorage().Open(path);
            foreach (var expectedSheet in root.GetProperty("Sheets").EnumerateArray())
            {
                var sheet = book.Sheets.Single(s => s.Name == expectedSheet.GetProperty("Sheet").GetString());
                Assert.Equal(expectedSheet.GetProperty("Rules").GetArrayLength(), sheet.Data.ConditionalFormats.Length);
                foreach (var cell in expectedSheet.GetProperty("Cells").EnumerateArray())
                {
                    var p = ScalarFormulaFixture.Position(cell.GetProperty("Address").GetString()!);
                    var style = ConditionalPresentation.GetStyle(book, sheet, p.Row, p.Column);
                    Assert.Equal(cell.GetProperty("FillColor").GetString(), style.FillColor);
                    Assert.Equal(cell.GetProperty("FontColor").GetString(), style.FontColor ?? "FF000000");
                    Assert.Equal(cell.GetProperty("Bold").GetBoolean(), style.Bold); Assert.Equal(cell.GetProperty("Italic").GetBoolean(), style.Italic);
                    Assert.Equal(cell.GetProperty("Strike").GetBoolean(), style.Strike);
                    Assert.Equal(cell.GetProperty("Text").GetString(), CellDisplayFormatter.Format(sheet.GetCell(p.Row, p.Column)?.Value ?? CellValue.Blank, style, book.DateSystem));
                }
            }
            Assert.Empty(ConditionalPresentation.Diagnostics(book));
            // Native filter evidence: conditional " units" changes display, while value filters use the base number format.
            var numbers = book.Sheets.Single(s => s.Name == "Numbers");
            foreach (var filter in root.GetProperty("FilterChecks").EnumerateArray())
            {
                var data = numbers.Data with { Filter = new(new(0, 0, 4, 0), [new(0, new(values: [filter.GetProperty("Criterion").GetString()!]))]) };
                var hidden = WorksheetFilterEngine.Evaluate(numbers.WithData(data), book.Styles, book.DateSystem, book.Revision);
                Assert.Equal(filter.GetProperty("VisibleRows").EnumerateArray().Select(v => v.GetInt32()), Enumerable.Range(1, 4).Where(r => !hidden.Contains(r)).Select(r => r + 1));
            }
        }
        finally { File.Delete(path); }
    }
}
