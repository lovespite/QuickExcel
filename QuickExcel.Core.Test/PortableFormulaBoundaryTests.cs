using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class PortableFormulaBoundaryTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "QuickExcel-portable-" + Guid.NewGuid().ToString("N"));
    public PortableFormulaBoundaryTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);

    [Fact]
    public void ExpandedNamesOverPrintLimitFreezeTheirClosureAndPreserveSource()
    {
        var text = new string('a', 4100);
        var sheet = Sheet("Long name", (0, 0, F("=LongText&LongText")), (0, 1, F("=LEN(A1)")), (0, 2, F("=2+3")));
        var book = Book(sheet).WithDefinedNames([new("LongText", "=\"" + text + "\"")]);
        using var session = new DocumentSession(book); var calculator = new WorkbookCalculator(new NoScripts());
        Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        var before = session.Snapshot; var storage = new TypedWorkbookStorage(); var report = storage.InspectPortableExport(before);
        Assert.Equal(2, report.FrozenFormulas); Assert.Equal(1, report.RetainedFormulas); Assert.Equal(0, report.StaleResults);
        var path = Path.Combine(_directory, "long.xlsx"); storage.ExportPortable(path, before);
        var saved = storage.Open(path); Assert.False(saved.FileState!.FormulaReadOnly); Assert.Empty(saved.DefinedNames);
        Assert.Equal(CellContentKind.Literal, saved.Sheets[0].GetCell(0, 0)!.Content.Kind);
        Assert.Equal(CellValue.FromText(text + text), saved.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(8200), saved.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Equal(CellContentKind.Formula, saved.Sheets[0].GetCell(0, 2)!.Content.Kind);
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        Assert.False(FormulaSyntax.TryPrintPortable(FormulaBinding.Bind(FormulaSyntax.Parse("=LongText&LongText"), book, sheet.Id), out _));
    }

    [Fact]
    public void IllegalArityAndItsDependentsNeverEscapeAsInvalidExcelFormulas()
    {
        var sheet = Sheet("S", (0, 0, F("=IF(1)")), (0, 1, F("=IFERROR(A1,9)")), (0, 2, F("=CONCAT(\"a\",\"b\")")));
        using var session = new DocumentSession(Book(sheet));
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        var storage = new TypedWorkbookStorage(); var report = storage.InspectPortableExport(session.Snapshot);
        Assert.Equal(2, report.FrozenFormulas); Assert.Equal(1, report.RetainedFormulas);
        var invalid = Path.Combine(_directory, "invalid.xlsx"); File.WriteAllText(invalid, "original");
        Assert.Throws<InvalidOperationException>(() => storage.Save(invalid, session.Snapshot));
        Assert.Throws<InvalidOperationException>(() => storage.Save(invalid, session.Snapshot, discardUnsupported: true));
        Assert.Equal("original", File.ReadAllText(invalid));
        Assert.Throws<InvalidOperationException>(() => storage.Save(invalid, Book(Sheet("Names", (0, 0, F("=BadName")))).WithDefinedNames([new("BadName", "=IF(1)")])));
        var path = Path.Combine(_directory, "arity.xlsx"); storage.ExportPortable(path, session.Snapshot);
        var saved = storage.Open(path); Assert.Equal(CellValue.FromError(CellErrorCode.Value), saved.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(9), saved.Sheets[0].GetCell(0, 1)!.Value);
        using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry("xl/workbook.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Assert.Equal("1", (string?)XDocument.Load(stream).Root!.Element(ns + "calcPr")!.Attribute("fullCalcOnLoad"));
    }

    [Fact]
    public void AggregateLazyBranchesInvokeSelectedWorkerUdfExactlyOnce()
    {
        var sheet = Sheet("S", (0, 0, F("=SUM(IF(TRUE,NEXT(),FAIL()))")), (1, 0, F("=COUNT(IF(FALSE,FAIL(),NEXT()))")),
            (2, 0, F("=SUM(IFERROR(NEXT(),FAIL()))")), (3, 0, F("=IFNA(NEXT(),FAIL())")));
        using var session = new DocumentSession(Book(sheet).WithScriptSource("let n=0; host.functions.register('NEXT',()=>++n); host.functions.register('FAIL',()=>{throw new Error('unselected')});"));
        session.SetScriptsEnabled(true);
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new ScriptWorkerRuntimeFactory()).Calculate(session.BeginCalculation())));
        Assert.Equal(new[] { 1d, 1d, 3d, 4d }, Enumerable.Range(0, 4).Select(row => session.Snapshot.Sheets[0].GetCell(row, 0)!.Value.NumberValue));
        Assert.All(session.Snapshot.Sheets[0].Cells.Values, cell => Assert.Equal(CalculationState.Fresh, cell.Calculation!.State));
        Assert.Equal(4, new TypedWorkbookStorage().InspectPortableExport(session.Snapshot).FrozenFormulas);
    }
}
