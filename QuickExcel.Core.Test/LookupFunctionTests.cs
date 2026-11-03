using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class LookupFunctionTests
{
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ExcelLookupMatrixMatchesIncrementalEditUndoRedoAndRoundtrip(string extension)
    {
        using var specJson = ScalarFormulaFixture.Load("lookup-function-spec"); using var expectedJson = ScalarFormulaFixture.Load("excel-lookup-functions");
        var spec = specJson.RootElement; var expected = expectedJson.RootElement;
        using var session = new DocumentSession(ScalarFormulaFixture.Book(spec)); var calculator = new WorkbookCalculator(new NoScripts());
        void Check(string state) { ScalarFormulaFixture.Calculate(session, calculator); ScalarFormulaFixture.Compare(session.Snapshot, spec, expected.GetProperty(state)); }
        Check("Initial"); ScalarFormulaFixture.Edit(session, spec); Check("Edited");
        Assert.True(session.Undo()); Assert.False(session.CanUndo); Check("Initial"); Assert.True(session.Redo()); Check("Edited");
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); using var reopened = new DocumentSession(storage.Open(path));
            Assert.False(reopened.Snapshot.FileState?.FormulaReadOnly); ScalarFormulaFixture.Calculate(reopened, new WorkbookCalculator(new NoScripts()));
            ScalarFormulaFixture.Compare(reopened.Snapshot, spec, expected.GetProperty("Edited"));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void PrefixEncodingPreservesCaseWhitespaceStringsAndNamesAndNormalizesOnlyImplementedFunctions()
    {
        var formula = "=xlookup (1, 'XMATCH'!A1:A2, B1:B2, \"_xlfn.XLOOKUP(\") + 1";
        var encoded = ExcelFormulaNames.Encode(formula);
        Assert.StartsWith("=_xlfn.xlookup (", encoded); Assert.Contains("'XMATCH'!", encoded); Assert.Contains("\"_xlfn.XLOOKUP(\"", encoded);
        Assert.Equal(formula, ExcelFormulaNames.Decode(encoded)); Assert.Equal(formula, CellContent.FromFormula(encoded).FormulaText);
        Assert.Equal("=_xlfn.FUTURELOOKUP(A1)", ExcelFormulaNames.Decode("=_xlfn.FUTURELOOKUP(A1)"));
        Assert.Equal("XMATCH", Assert.IsType<FunctionNode>(FormulaSyntax.Parse("=_xlfn.xmatch(1,A1:A2)")).Name);
    }
    [Fact]
    public void ExtensionFunctionsAndDefinedNamesHaveRequiredPrefixesInXlsx()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (0, 1, N(5)), (0, 2, F("=xlookup (1,A1:A2,B1:B2)+1")), (1, 2, F("=Pick")));
        var book = Book(sheet).WithDefinedNames([new("Pick", "=_xlfn.XLOOKUP(1,S!$A$1:$A$2,S!$B$1:$B$2)")]);
        using var session = new DocumentSession(book); ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            Assert.Equal("=xlookup (1,A1:A2,B1:B2)+1", reopened.Sheets[0].GetCell(0, 2)!.Content.FormulaText); Assert.Equal(6, reopened.Sheets[0].GetCell(0, 2)!.Value.NumberValue);
            Assert.Equal(5, reopened.Sheets[0].GetCell(1, 2)!.Value.NumberValue); Assert.StartsWith("=XLOOKUP", reopened.DefinedNames[0].Formula);
            using var zip = ZipFile.OpenRead(path); using var cellStream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open(); var cells = XDocument.Load(cellStream);
            using var nameStream = zip.GetEntry("xl/workbook.xml")!.Open(); var names = XDocument.Load(nameStream); XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            Assert.Contains(cells.Descendants(ns + "f"), f => f.Value.StartsWith("_xlfn.xlookup (")); Assert.StartsWith("_xlfn.XLOOKUP", names.Descendants(ns + "definedName").Single().Value);
        }
        finally { File.Delete(path); }
    }
    [Theory]
    [InlineData("=XMATCH(2,A1:A7,0,2)")]
    [InlineData("=XMATCH(2,A1:A7,0,-1.5)")]
    [InlineData("=XMATCH(2,A1:A7,3)")]
    [InlineData("=XMATCH(2,A1:A7,0,B1)")]
    public void UnimplementedLookupModesImportReadOnlyAndKeepOriginalCache(string formula)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx"); var copy = path + ".copy.xlsx";
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(Sheet("S", (0, 4, F(formula, 123))))); var book = storage.Open(path);
            Assert.True(book.FileState!.FormulaReadOnly); using var session = new DocumentSession(book); ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
            Assert.Equal(CalculationState.Blocked, session.Snapshot.Sheets[0].GetCell(0, 4)!.Calculation!.State); Assert.Equal(123, session.Snapshot.Sheets[0].GetCell(0, 4)!.Value.NumberValue);
            storage.Save(copy, session.Snapshot); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(copy));
        }
        finally { File.Delete(path); File.Delete(copy); }
    }
}
