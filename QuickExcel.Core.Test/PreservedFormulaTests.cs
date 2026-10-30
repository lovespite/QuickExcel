using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Runtime;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class PreservedFormulaTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-preserved-formula-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public PreservedFormulaTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private string PathFor(string name) => Path.Combine(_folder, name);
    private static byte[] Part(string path, string name)
    { using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry(name)!.Open(); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void ReplacePart(string path, string name, XDocument xml)
    { using var zip = ZipFile.Open(path, ZipArchiveMode.Update); zip.GetEntry(name)!.Delete(); using var stream = zip.CreateEntry(name).Open(); xml.Save(stream); }
    private string Source(string kind)
    {
        var path = PathFor(kind + ".xlsx");
        var first = Sheet("Data", (0, 0, N(9)), (1, 0, N(4)), (0, 1, F("=SQRT(A1)", 3)), (1, 1, N(8)), (0, 2, F("=B1+1", 4)), (0, 3, F("=SUM(A1,2)", 11)));
        _storage.Save(path, Book(first, Sheet("Other", (0, 0, F("=Data!C1+1", 5)))));
        if (kind == "unknown") return path;
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/worksheets/sheet1.xml")));
        var formula = xml.Descendants(Main + "c").Single(c => (string?)c.Attribute("r") == "B1").Element(Main + "f")!;
        if (kind == "array") { formula.Value = "A1:A2*2"; formula.SetAttributeValue("t", "array"); formula.SetAttributeValue("ref", "B1:B2"); }
        else if (kind == "structured") formula.Value = "SUM(Table1[Amount])";
        else if (kind == "prefix") formula.Value = "_xlfn.XLOOKUP(A1,A1:A2,B1:B2)";
        else if (kind == "shared")
        {
            formula.Value = "SUM(Table1[Amount])"; formula.SetAttributeValue("t", "shared"); formula.SetAttributeValue("si", "7"); formula.SetAttributeValue("ref", "B1:B2");
            xml.Descendants(Main + "c").Single(c => (string?)c.Attribute("r") == "B2").AddFirst(new XElement(Main + "f", new XAttribute("t", "shared"), new XAttribute("si", 7)));
        }
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update); zip.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
        using var stream = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(stream); return path;
    }

    [Theory]
    [InlineData("array", PreservedFormulaKind.Array)]
    [InlineData("structured", PreservedFormulaKind.StructuredReference)]
    [InlineData("shared", PreservedFormulaKind.Shared)]
    public void UnsupportedSemanticsOpenReadOnlyAndPreserveOriginalPackageAfterNavigationAndCalculation(string kind, PreservedFormulaKind expected)
    {
        var path = Source(kind); var bytes = File.ReadAllBytes(path); var opened = _storage.Open(path);
        Assert.True(opened.FileState!.FormulaReadOnly); Assert.Contains(opened.FileState.PreservedFormulas, p => p.Kind == expected);
        using var session = new DocumentSession(opened);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        Assert.All(result.Cells, pair => { Assert.Equal(CalculationState.Blocked, pair.Value.State); Assert.Equal(opened.GetSheet(pair.Key.SheetId).GetCell(pair.Key.Row, pair.Key.Column)!.Calculation?.Value, pair.Value.Value); });
        Assert.True(session.TryCommitCalculation(result));
        session.UpdateNavigation(opened.Sheets[1].Id, new(20, 4), new(15, 2)); Assert.False(session.IsDirty); Assert.False(session.CanUndo);
        Assert.Empty(TypedWorkbookStorage.SaveIssues(session.Snapshot));
        Assert.True(TypedWorkbookStorage.InspectCompatibility(session.Snapshot).OriginalPackageUnchanged);
        var copy = PathFor("copy.xlsx"); _storage.Save(copy, session.Snapshot); Assert.Equal(bytes, File.ReadAllBytes(copy));
        var container = PathFor("copy.xlsq"); _storage.Save(container, session.Snapshot); Assert.Equal(bytes, Part(container, "workbook.xlsx"));
        Assert.Contains("preserved-formulas", Encoding.UTF8.GetString(Part(container, "manifest.json")));
        Assert.True(_storage.Open(container).FileState!.FormulaReadOnly); Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public void ReadOnlyPolicyRejectsEditingForgedCalculationAndLossyFormulaCopiesAtomically()
    {
        var path = Source("array"); var book = _storage.Open(path); var id = book.Sheets[0].Id; using var session = new DocumentSession(book);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))])));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(Bold: true))));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1)));
        Assert.Throws<InvalidOperationException>(() => session.SetScriptsEnabled(true));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo); Assert.False(session.IsDirty);
        var request = session.BeginCalculation(); var stamp = request.Stamp;
        var forged = new CalculationResult(stamp, new Dictionary<CellAddress, CellCalculation> { [new(id, 0, 1)] = new(CellValue.FromNumber(42), CalculationState.Fresh, stamp.InputRevision, stamp.CalculationGeneration, stamp.ScriptGeneration) });
        Assert.Throws<InvalidOperationException>(() => session.TryCommitCalculation(forged)); Assert.Same(book, session.Snapshot);
        Assert.Throws<InvalidOperationException>(() => _storage.Save(PathFor("loss.xlsx"), book, discardUnsupported: true));
        Assert.Throws<InvalidOperationException>(() => _storage.Save(PathFor("changed.xlsx"), book.WithSheets(book.Sheets.Select(s => s.Id == id ? s.WithCells(s.Cells.SetItem(new(0, 0), N(42))) : s))));
        Assert.Throws<InvalidOperationException>(() => _storage.ExportPortable(PathFor("portable.xlsx"), book, true, true));
        Assert.Throws<InvalidOperationException>(() => _storage.Export(PathFor("values.xlsx"), book, false, true));
        _storage.Export(PathFor("values.xlsx"), book, true, true);
        var values = _storage.Open(PathFor("values.xlsx")); Assert.False(values.FileState!.FormulaReadOnly);
        Assert.All(values.Sheets.SelectMany(s => s.Cells.Values), c => Assert.Equal(CellContentKind.Literal, c.Content.Kind));
        Assert.Equal(CellValue.FromNumber(8), values.Sheets[0].GetCell(1, 1)!.Value);
    }

    [Fact]
    public void UnknownFunctionAndDependentsKeepOldCachesWhileIndependentFormulasRecalculate()
    {
        var path = Source("unknown"); var book = _storage.Open(path); Assert.False(book.FileState!.FormulaReadOnly);
        using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(16)))]));
        var beforeStructure = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new RenameSheetCommand(id, "Renamed"))); Assert.Same(beforeStructure, session.Snapshot);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation()); Assert.True(session.TryCommitCalculation(result));
        Assert.Equal(CalculationState.Blocked, result.Cells[new(id, 0, 1)].State); Assert.Equal(CellValue.FromNumber(3), result.Cells[new(id, 0, 1)].Value);
        Assert.Equal(CalculationState.Blocked, result.Cells[new(id, 0, 2)].State); Assert.Equal(CellValue.FromNumber(4), result.Cells[new(id, 0, 2)].Value);
        Assert.Equal(CalculationState.Blocked, result.Cells[new(book.Sheets[1].Id, 0, 0)].State);
        Assert.Equal(CalculationState.Fresh, result.Cells[new(id, 0, 3)].State); Assert.Equal(CellValue.FromNumber(18), result.Cells[new(id, 0, 3)].Value);
        _storage.Save(PathFor("edited.xlsx"), session.Snapshot); var saved = _storage.Open(PathFor("edited.xlsx"));
        Assert.Equal("=SQRT(A1)", saved.Sheets[0].GetCell(0, 1)!.Content.FormulaText); Assert.Equal(CellValue.FromNumber(3), saved.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Equal(CalculationState.Stale, saved.Sheets[0].GetCell(0, 1)!.Calculation!.State);
    }
    [Fact]
    public void VersionPrefixedFunctionsRemainVerbatimAndAreReportedSeparately()
    {
        var book = _storage.Open(Source("prefix")); Assert.False(book.FileState!.FormulaReadOnly);
        Assert.Contains(TypedWorkbookStorage.InspectCompatibility(book).Diagnostics, d => d.Feature == "formula/version-prefix");
        Assert.StartsWith("=_xlfn.XLOOKUP", book.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
    }

    [Fact]
    public void MissingArrayFollowerCachePreventsEvenExplicitValueExport()
    {
        var path = Source("array"); var xml = XDocument.Load(new MemoryStream(Part(path, "xl/worksheets/sheet1.xml")));
        xml.Descendants(Main + "c").Single(c => (string?)c.Attribute("r") == "B2").Element(Main + "v")!.Remove();
        ReplacePart(path, "xl/worksheets/sheet1.xml", xml); var opened = _storage.Open(path);
        Assert.False(Assert.Single(opened.FileState!.PreservedFormulas).HasCompleteCache);
        Assert.Throws<InvalidOperationException>(() => _storage.Export(PathFor("missing.xlsx"), opened, true, true));
        Assert.False(File.Exists(PathFor("missing.xlsx")));
    }

    [Fact]
    public void UnknownFunctionInAPartialCycleCannotReplaceItsSourceCacheWithANewCycleError()
    {
        var book = _storage.Open(Source("prefix")); using var session = new DocumentSession(book);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation()); var id = book.Sheets[0].Id;
        Assert.Equal(CalculationState.Blocked, result.Cells[new(id, 0, 1)].State); Assert.Equal(CellValue.FromNumber(3), result.Cells[new(id, 0, 1)].Value);
        Assert.Equal(CalculationState.Blocked, result.Cells[new(id, 0, 2)].State); Assert.Equal(CellValue.FromNumber(4), result.Cells[new(id, 0, 2)].Value);
    }

    [Fact]
    public void EnablingScriptsOnlyUnblocksFunctionsActuallyRegisteredInThatRuntime()
    {
        var source = "quickExcel.functions.register(\"DOUBLE\", value => value * 2);";
        var sheet = Sheet("Data", (0, 0, N(9)), (0, 1, F("=SQRT(A1)", 3)), (0, 2, F("=B1+1", 4)),
            (0, 3, F("=DOUBLE(A1)", 0)), (0, 4, F("=D1+1", 0)));
        var book = Book(sheet).WithScriptSource(source);
        var result = new WorkbookCalculator(new QuickJsRuntimeFactory()).Calculate(Request(book, true));
        Assert.Equal(CalculationState.Blocked, result.Cells[new(sheet.Id, 0, 1)].State);
        Assert.Equal(CellValue.FromNumber(3), result.Cells[new(sheet.Id, 0, 1)].Value);
        Assert.Equal(CalculationState.Blocked, result.Cells[new(sheet.Id, 0, 2)].State);
        Assert.Equal(CalculationState.Fresh, result.Cells[new(sheet.Id, 0, 3)].State);
        Assert.Equal(CellValue.FromNumber(19), result.Cells[new(sheet.Id, 0, 4)].Value);
    }

    [Fact]
    public void UnsupportedDefinedNameDoesNotTurnItsSourceResultIntoAFreshNameError()
    {
        var path = Source("unknown");
        var workbook = XDocument.Load(new MemoryStream(Part(path, "xl/workbook.xml")));
        workbook.Root!.Add(new XElement(Main + "definedNames", new XElement(Main + "definedName", new XAttribute("name", "Amounts"), "Table1[Amount]")));
        ReplacePart(path, "xl/workbook.xml", workbook);
        var xml = XDocument.Load(new MemoryStream(Part(path, "xl/worksheets/sheet1.xml")));
        xml.Descendants(Main + "f").First().Value = "SUM(Amounts)";
        ReplacePart(path, "xl/worksheets/sheet1.xml", xml);
        var opened = _storage.Open(path); Assert.True(opened.FileState!.FormulaReadOnly);
        Assert.Contains(opened.FileState.PreservedFormulas, f => f.Kind == PreservedFormulaKind.DefinedName && f.Cell == "Amounts");
        var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(opened));
        Assert.Equal(CalculationState.Blocked, result.Cells[new(opened.Sheets[0].Id, 0, 1)].State);
        Assert.Equal(CellValue.FromNumber(3), result.Cells[new(opened.Sheets[0].Id, 0, 1)].Value);
    }
}
