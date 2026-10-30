using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ConditionalFunctionTests
{
    private static JsonDocument Fixture() => JsonDocument.Parse(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-conditional-functions.json")));
    private static CellValue Value(JsonElement item) => item.GetProperty("Kind").GetString() switch
    {
        "Number" => CellValue.FromNumber(item.GetProperty("Value").GetDouble()),
        "Text" => CellValue.FromText(item.GetProperty("Value").GetString()!),
        "Boolean" => CellValue.FromBoolean(item.GetProperty("Value").GetBoolean()),
        "Error" => ((ValueNode)FormulaSyntax.Parse("=" + item.GetProperty("Value").GetString())).Value,
        _ => CellValue.Blank
    };
    private static WorkbookSnapshot BookFrom(JsonElement fixture)
    {
        var cells = fixture.GetProperty("Inputs").EnumerateArray().Select(item =>
        {
            var address = (ReferenceNode)FormulaSyntax.Parse("=" + item.GetProperty("Address").GetString());
            var content = item.GetProperty("Kind").GetString() == "Formula" ? CellContent.FromFormula(item.GetProperty("Value").GetString()!) : CellContent.FromLiteral(Value(item));
            return new KeyValuePair<CellPosition, CellSnapshot>(new(address.Row, address.Column), new(content));
        });
        var data = new WorksheetSnapshot(SheetId.New(), "Data", cells);
        var results = new WorksheetSnapshot(SheetId.New(), "Results", fixture.GetProperty("Cases").EnumerateArray().Select((item, i) =>
            new KeyValuePair<CellPosition, CellSnapshot>(new(i, 0), F(item.GetProperty("Formula").GetString()!))));
        return Book(data, results);
    }
    private static void Compare(WorkbookSnapshot book, JsonElement fixture, bool edited)
    {
        var differences = new List<string>(); var cases = fixture.GetProperty("Cases").EnumerateArray().ToArray(); var sheet = book.Sheets.Single(s => s.Name == "Results");
        for (var i = 0; i < cases.Length; i++)
        {
            var expected = Value(cases[i].GetProperty(edited ? "Edited" : "Initial")); var actual = sheet.GetCell(i, 0)!.Value;
            var same = expected.Kind == actual.Kind && (actual.Kind == CellValueKind.Number
                ? Math.Abs(expected.NumberValue - actual.NumberValue) <= 1e-12 * Math.Max(1, Math.Abs(expected.NumberValue)) : expected == actual);
            if (!same) differences.Add($"{cases[i].GetProperty("Id").GetString()} {cases[i].GetProperty("Formula").GetString()}: expected {expected}, actual {actual}");
        }
        Assert.True(differences.Count == 0, $"{differences.Count} mismatches\n" + string.Join("\n", differences.Take(35)));
    }
    private static CalculationResult Calculate(DocumentSession session, WorkbookCalculator calculator)
    { var result = calculator.Calculate(session.BeginCalculation()); Assert.True(session.TryCommitCalculation(result)); return result; }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ExcelFixedMatrixMatchesIncrementalEditUndoRedoAndRoundtrip(string extension)
    {
        using var json = Fixture(); var fixture = json.RootElement; Assert.Equal(1706, fixture.GetProperty("Cases").GetArrayLength());
        using var session = new DocumentSession(BookFrom(fixture)); var calculator = new WorkbookCalculator(new NoScripts());
        Calculate(session, calculator); Compare(session.Snapshot, fixture, false);
        var data = session.Snapshot.Sheets[0];
        session.Execute(new SetCellsCommand([new(new(data.Id, 0, 0), N(7).Content), new(new(data.Id, 2, 1), N(300).Content)]));
        var edited = Calculate(session, calculator); Compare(session.Snapshot, fixture, true); Assert.True(Assert.IsType<CalculationStatistics>(edited.Statistics).ReusedCount > 0);
        session.Undo(); Calculate(session, calculator); Compare(session.Snapshot, fixture, false); Assert.False(session.CanUndo);
        session.Redo(); Calculate(session, calculator); Compare(session.Snapshot, fixture, true);
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); using var reopened = new DocumentSession(storage.Open(path));
            Assert.False(reopened.Snapshot.FileState?.FormulaReadOnly); Calculate(reopened, new WorkbookCalculator(new NoScripts())); Compare(reopened.Snapshot, fixture, true);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ProjectedResultDependenciesIncludeOutsideWrittenAnchorAndExcludeUnusedTail()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (1, 0, N(1)), (2, 0, N(1)), (0, 1, N(5)), (1, 1, N(6)), (2, 1, F("=A1*10")),
            (0, 3, F("=SUMIF(A1:A3,1,B1)")), (4, 1, F("=SUMIF(A1,1,B1:B10)")), (5, 1, F("=AVERAGEIF(A1:A3,1,B1)")));
        using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts()); Calculate(session, calculator);
        Assert.Equal(21, session.Snapshot.Sheets[0].GetCell(0, 3)!.Value.NumberValue); Assert.Equal(5, session.Snapshot.Sheets[0].GetCell(4, 1)!.Value.NumberValue);
        Assert.Equal(7, session.Snapshot.Sheets[0].GetCell(5, 1)!.Value.NumberValue);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 2, 1), N(40).Content)])); Calculate(session, calculator);
        Assert.Equal(51, session.Snapshot.Sheets[0].GetCell(0, 3)!.Value.NumberValue); Assert.Equal(17, session.Snapshot.Sheets[0].GetCell(5, 1)!.Value.NumberValue);
    }

    [Fact]
    public void NamedRangesAndMaximumCriterionPairsUseTheSameBindingAndShapeRules()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (1, 0, N(2)), (0, 1, N(10)), (1, 1, N(20)),
            (0, 3, F("=SUMIF(Items,\">1\",Values)")), (1, 3, F("=COUNTIFS(" + string.Join(",", Enumerable.Repeat("A1:A2,\">0\"", 127)) + ")")),
            (2, 3, F("=COUNTIFS(" + string.Join(",", Enumerable.Repeat("A1:A2,\">0\"", 128)) + ")")));
        var book = Book(sheet).WithDefinedNames([new("Items", "=S!$A$1:$A$2"), new("Values", "=S!$B$1:$B$2")]);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(book));
        Assert.Equal(CellValue.FromNumber(20), C(result, sheet, 0, 3).Value);
        Assert.Equal(CellValue.FromNumber(2), C(result, sheet, 1, 3).Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), C(result, sheet, 2, 3).Value);
    }

    [Theory]
    [InlineData("=COUNTIF(1,1)")]
    [InlineData("=COUNTBLANK(0)")]
    [InlineData("=SUMIF(A1:A2,1,5)")]
    [InlineData("=AVERAGEIF(A1:A2,1,5)")]
    [InlineData("=COUNTIFS(A1:A2,1,A1:A2)")]
    [InlineData("=SUMIFS(B1:B2,A1:A2)")]
    [InlineData("=AVERAGEIFS(B1:B2,A1:A2,1,A1:A2)")]
    public void InvalidArityAndNonReferenceRangesAreCellErrors(string formula)
    {
        var sheet = Sheet("S", (0, 4, F(formula))); var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet)));
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), C(result, sheet, 0, 4).Value);
    }

    [Fact]
    public void ArrayCriterionImportRetainsSourceCacheAndOriginalPackage()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx"); var copy = path + ".copy.xlsx";
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(Sheet("S", (0, 0, F("=COUNTIF(B1:B2,1)", 7)))));
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"; xml.Descendants(ns + "f").Single().Value = "COUNTIF(B1:B2,C1:C2)";
                entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
            }
            var book = storage.Open(path); Assert.True(book.FileState!.FormulaReadOnly);
            using var session = new DocumentSession(book); Calculate(session, new WorkbookCalculator(new NoScripts()));
            Assert.Equal(CalculationState.Blocked, session.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
            Assert.Equal(7, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
            storage.Save(copy, session.Snapshot); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(copy));
        }
        finally { File.Delete(path); File.Delete(copy); }
    }
}
