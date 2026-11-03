using System.IO.Compression;
using System.Text.Json;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M3StorageTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-M3-" + Guid.NewGuid().ToString("N"));
    public M3StorageTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private string PathFor(string name) => Path.Combine(_folder, name);

    [Fact]
    public void NamesStableScopeAndTypedCrossSheetCachesSurviveXlsqWhileReopenRevokesFreshness()
    {
        var data = Sheet("Data", (0, 0, N(7)));
        var report = Sheet("Report", (0, 0, F("=Rate+Local")));
        var book = Book(data, report).WithDefinedNames([new("Rate", "='Data'!$A$1"), new("Local", "=3", report.Id)]);
        using var document = new WorkbookDocument(book, new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        document.Calculate();
        var path = PathFor("names.xlsq"); document.Save(path);
        var reopened = new TypedWorkbookStorage().Open(path);
        Assert.Equal(book.DefinedNames.OrderBy(n => n.Name), reopened.DefinedNames.OrderBy(n => n.Name));
        Assert.Equal(report.Id, reopened.Sheets[1].Id);
        Assert.Equal(CellValue.FromNumber(10), reopened.Sheets[1].GetCell(0, 0)!.Value);
        Assert.Equal(CalculationState.Stale, reopened.Sheets[1].GetCell(0, 0)!.Calculation!.State);
        using var zip = ZipFile.OpenRead(path);
        using var json = JsonDocument.Parse(zip.GetEntry("state/document.json")!.Open());
        Assert.Equal("fresh", json.RootElement.GetProperty("calculation").GetProperty("state").GetString());
        Assert.Equal("QuickExcel-M3", json.RootElement.GetProperty("calculation").GetProperty("engineVersion").GetString());
    }

    [Fact]
    public void PortableExportFreezesScriptCrossSheetTransitiveClosureButRetainsIndependentFormulas()
    {
        var a = Sheet("A", (0, 0, F("=DOUBLE(7)")), (0, 1, F("=4+5")), (0, 2, N(4)));
        var b = Sheet("B", (0, 0, F("=A!A1*2")), (0, 1, F("=SUM(A!A1:A2)")), (0, 2, F("=A!B1+1")));
        var c = Sheet("C", (0, 0, F("=B!A1+1")), (0, 1, F("=SUM(Independent)")));
        var book = Book(c, b, a).WithScriptSource("quickExcel.functions.register('DOUBLE', n => n*2);")
            .WithDefinedNames([new("Independent", "='A'!$B$1:$C$1")]);
        using var doc = new WorkbookDocument(book, new(new TypedWorkbookStorage(), new WorkbookCalculator(new JscRuntimeFactory())));
        doc.EnableScriptsAndCalculate();
        var report = doc.InspectPortableExport();
        Assert.Equal(4, report.FrozenFormulas); Assert.Equal(3, report.RetainedFormulas);
        var path = PathFor("portable.xlsx"); doc.ExportPortableXlsx(path);
        var saved = new TypedWorkbookStorage().Open(path);
        Assert.Equal(CellContentKind.Literal, saved.Sheets[0].GetCell(0, 0)!.Content.Kind);
        Assert.Equal(CellValue.FromNumber(29), saved.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellContentKind.Formula, saved.Sheets[0].GetCell(0, 1)!.Content.Kind);
        Assert.Equal(CellValue.FromNumber(13), saved.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Empty(saved.DefinedNames);
        using var zip = ZipFile.OpenRead(path);
        Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("scripts"));
        var calculated = new WorkbookCalculator(new NoScripts()).Calculate(Request(saved));
        Assert.Equal(CellValue.FromNumber(13), C(calculated, saved.Sheets[0], 0, 1).Value);
        Assert.True(doc.ScriptsEnabled);
    }

    [Fact]
    public void StaleOrMissingFrozenResultsRequireExplicitChoiceAndFailedExportPreservesFile()
    {
        var sheet = Sheet("S", (0, 0, F("=UNKNOWN(1)", 18)));
        var book = Book(sheet);
        var storage = new TypedWorkbookStorage();
        var path = PathFor("target.xlsx"); File.WriteAllText(path, "original");
        Assert.Throws<InvalidOperationException>(() => storage.ExportPortable(path, book));
        Assert.Equal("original", File.ReadAllText(path));
        storage.ExportPortable(path, book, allowStaleCaches: true);
        Assert.Equal(CellValue.FromNumber(18), storage.Open(path).Sheets[0].GetCell(0, 0)!.Value);
        var missing = Book(Sheet("S", (0, 0, F("=UNKNOWN(1)"))));
        Assert.Throws<InvalidOperationException>(() => storage.ExportPortable(PathFor("missing.xlsx"), missing, true));
        Assert.False(File.Exists(PathFor("missing.xlsx")));
    }

    [Fact]
    public void PortableStandardFormulasCanOmitMissingOrStaleCachesForExcelToRecalculate()
    {
        var sheet = Sheet("S", (0, 0, F("=2+3", 99)), (0, 1, F("=A1*2")));
        var storage = new TypedWorkbookStorage(); var path = PathFor("standard.xlsx");
        storage.ExportPortable(path, Book(sheet));
        var opened = storage.Open(path);
        Assert.Null(opened.Sheets[0].GetCell(0, 0)!.Calculation!.Value);
        var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(opened));
        Assert.Equal(CellValue.FromNumber(10), C(result, opened.Sheets[0], 0, 1).Value);
    }

    [Fact]
    public void ScriptInDefinedNameIsPartOfFrozenClosure()
    {
        var sheet = Sheet("S", (0, 0, F("=Rate+1")), (0, 1, F("=A1*2")));
        var book = Book(sheet).WithScriptSource("quickExcel.functions.register('NUMBER', ()=>9)")
            .WithDefinedNames([new("Rate", "=NUMBER()")]);
        using var doc = new WorkbookDocument(book, new(new TypedWorkbookStorage(), new WorkbookCalculator(new JscRuntimeFactory())));
        doc.EnableScriptsAndCalculate();
        Assert.Equal(2, doc.InspectPortableExport().FrozenFormulas);
        var path = PathFor("name-script.xlsx"); doc.ExportPortableXlsx(path);
        var saved = new TypedWorkbookStorage().Open(path);
        Assert.Equal(CellValue.FromNumber(20), saved.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Empty(saved.DefinedNames);
    }

    [Fact]
    public void ValuesOnlyExportDropsDefinitionsThatCouldContainScripts()
    {
        var sheet = Sheet("S", (0, 0, N(4)));
        var book = Book(sheet).WithDefinedNames([new("Unused", "=JS.CODE()")]);
        var path = PathFor("values.xlsx");
        var storage = new TypedWorkbookStorage();
        storage.Export(path, book);
        Assert.Empty(storage.Open(path).DefinedNames);
        using var zip = ZipFile.OpenRead(path);
        using var reader = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open());
        Assert.DoesNotContain("JS.CODE", reader.ReadToEnd());
    }

    [Fact]
    public void ChangedInputPersistsStaleAggregateAndIndependentSaveDoesNotRunScripts()
    {
        var sheet = Sheet("S", (0, 0, N(3)), (0, 1, F("=A1*2")));
        using var doc = new WorkbookDocument(Book(sheet), new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts())));
        doc.Calculate(); doc.SetCell(0, 0, 0, "7");
        var path = PathFor("stale.xlsq"); doc.Save(path);
        using var zip = ZipFile.OpenRead(path); using var json = JsonDocument.Parse(zip.GetEntry("state/document.json")!.Open());
        Assert.Equal("stale", json.RootElement.GetProperty("calculation").GetProperty("state").GetString());
        var opened = new TypedWorkbookStorage().Open(path);
        Assert.Equal(CellValue.FromNumber(6), opened.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Equal(CalculationState.Stale, opened.Sheets[0].GetCell(0, 1)!.Calculation!.State);
    }
}
