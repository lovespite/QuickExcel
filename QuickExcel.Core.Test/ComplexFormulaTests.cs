using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ComplexFormulaTests
{
    [Fact]
    public void NativeLetAliasesAndStructuralEditsPreserveLocalOwnership()
    {
        const string formula = "=LET( x, 'x'!A1, LET(x,x+1,x) + x + LEN(\"_xlpm.x\"))";
        const string encoded = "=_xlfn.LET( _xlpm.x, 'x'!A1, _xlfn.LET(_xlpm.x,_xlpm.x+1,_xlpm.x) + _xlpm.x + LEN(\"_xlpm.x\"))";
        Assert.Equal(encoded, ExcelFormulaNames.Encode(formula));
        Assert.Equal(formula, ExcelFormulaNames.Decode(encoded));
        var source = Sheet("Data", (0, 0, F("=LET(local_value,2,local_value+global_value)")));
        var target = Sheet("Target");
        using var session = new DocumentSession(Book(source, target).WithDefinedNames([new DefinedName("global_value", "=3")]));
        Assert.Throws<NotSupportedException>(() => session.Execute(new RenameDefinedNameCommand("global_value", null, "local_value")));
        Assert.False(session.CanUndo);
        session.Execute(new RenameDefinedNameCommand("global_value", null, "renamed_value"));
        session.Execute(new MoveCellsCommand(source.Id, new(0, 0), new(0, 0), target.Id, new(2, 2)));
        ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
        var cell = session.Snapshot.GetSheet(target.Id).GetCell(2, 2)!;
        Assert.Equal(CellValue.FromNumber(5), cell.Value);
        Assert.Contains("LET(local_value,2,", cell.Content.FormulaText);
        Assert.DoesNotContain("'Data'!local_value", cell.Content.FormulaText);
        var chain = "=LET(" + string.Join(',', Enumerable.Range(0, 100).Select(i => $"local_{i}," + (i == 0 ? "1" : $"local_{i - 1}+1"))) + ",local_99)";
        Assert.Throws<FormatException>(() => FormulaBinding.Bind(FormulaSyntax.Parse(chain), session.Snapshot, target.Id));
    }

    [Fact]
    public void LocalNamesKeepLexicalScopeAndScriptValuesAreEvaluatedOnce()
    {
        var sheet = Sheet("S", (0, 0, F("=LET(x,NEXT(),x+x)")),
            (1, 0, F("=LET(x,NEXT(),LET(y,x+x,y+x))")), (2, 0, F("=LET(x,NEXT(),5)")), (3, 0, F("=NEXT()")));
        var book = Book(sheet).WithScriptSource("let n=0;quickExcel.functions.register('NEXT',()=>++n);");
        var result = new WorkbookCalculator(new QuickExcel.Runtime.JscRuntimeFactory()).Calculate(Request(book) with { ScriptsEnabled = true });
        Assert.Equal(new double[] { 2, 6, 5, 3 }, Enumerable.Range(0, 4).Select(row => result.Cells[new(sheet.Id, row, 0)].Value!.Value.NumberValue));
        var parsed = FormulaSyntax.Parse("=LET(x,global_name,LET(y,x,y)+x+global_name)");
        var mapped = FormulaSyntax.Map(parsed, n => n is NameNode name ? name with { Sheet = "Other" } : n);
        Assert.Equal("LET(x,'Other'!global_name,((LET(y,x,y)+x)+'Other'!global_name))", FormulaSyntax.Print(mapped));
        Assert.Equal(new[] { "global_name", "global_name" }, FormulaSyntax.Walk(parsed).OfType<NameNode>().Select(n => n.Name));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void IndependentExcelMatrixMatchesIncrementalEditUndoAndRoundtrip(string extension)
    {
        using var specJson = ScalarFormulaFixture.Load("complex-function-spec");
        using var oracle = ScalarFormulaFixture.Load("excel-complex-functions");
        var spec = specJson.RootElement;
        using var session = new DocumentSession(ScalarFormulaFixture.Book(spec));
        var calculator = new WorkbookCalculator(new NoScripts());
        void Check(string state) { ScalarFormulaFixture.Calculate(session, calculator); ScalarFormulaFixture.Compare(session.Snapshot, spec, oracle.RootElement.GetProperty(state)); }
        Check("Initial"); ScalarFormulaFixture.Edit(session, spec); Check("Edited");
        Assert.True(session.Undo()); Check("Initial"); Assert.True(session.Redo()); Check("Edited");
        var folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../e5/complex", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(folder); var path = Path.Combine(folder, "complex." + extension);
        var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot);
        using var reopened = new DocumentSession(storage.Open(path));
        Assert.False(reopened.Snapshot.FileState?.FormulaReadOnly);
        ScalarFormulaFixture.Calculate(reopened, new WorkbookCalculator(new NoScripts()));
        ScalarFormulaFixture.Compare(reopened.Snapshot, spec, oracle.RootElement.GetProperty("Edited"));
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(folder)!, "latest-" + extension + ".txt"), path);
    }

    [Theory]
    [InlineData("=OFFSET(A1,A2,0)")]
    [InlineData("=INDIRECT(A2)")]
    [InlineData("=INDIRECT(\"[other.xlsx]S!A1\")")]
    [InlineData("=LAMBDA(x,x+1)(2)")]
    [InlineData("=SUM(S1:S3!A1)")]
    [InlineData("=INDIRECT(\"XFD1048576\")")]
    [InlineData("=INDIRECT(\"A1:A3\")")]
    [InlineData("=OFFSET(A1,0,0,3,1)")]
    [InlineData("=INDIRECT(\"1:3\")")]
    public void UnmodeledReferencesNeverUseOldCacheAsFreshResult(string formula)
    {
        var folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../e5/complex", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(folder); var path = Path.Combine(folder, "guard.xlsx");
        var cellXml = new System.Xml.Linq.XElement("row", new System.Xml.Linq.XAttribute("r", 1),
            new System.Xml.Linq.XElement("c", new System.Xml.Linq.XAttribute("r", "D1"),
                new System.Xml.Linq.XElement("f", formula[1..]), new System.Xml.Linq.XElement("v", 123)));
        File.WriteAllBytes(path, Fixtures.WorkbookPackageFixture.Xlsx(cellXml.ToString()));
        var storage = new TypedWorkbookStorage();
        using var session = new DocumentSession(storage.Open(path));
        Assert.True(session.Snapshot.FileState!.FormulaReadOnly);
        ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
        var cell = session.Snapshot.Sheets[0].GetCell(0, 3)!;
        Assert.Equal(CalculationState.Blocked, cell.Calculation!.State);
        Assert.Equal(CellValue.FromNumber(123), cell.Value);
        var copy = Path.Combine(folder, "preserved.xlsx"); storage.Save(copy, session.Snapshot);
        Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(copy));
    }

    [Fact]
    public void StaticReferencesPlanTargetAndKeepOriginalPortableSyntax()
    {
        var sheet = Sheet("S", (0, 0, F("=OFFSET(A1,0,1)")), (0, 1, F("=2+3")),
            (1, 0, F("=INDIRECT(\"B1\")")), (2, 0, F("=LET(x,2,x+LET(y,x+1,y))")));
        var book = Book(sheet); var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(book));
        Assert.All(new[] { 0, 1, 2 }, row => Assert.Equal(CellValue.FromNumber(5), result.Cells[new(sheet.Id, row, 0)].Value));
        var index = new WorkbookDependencyIndex(book);
        Assert.DoesNotContain(new CellAddress(sheet.Id, 0, 0), index.CycleMembers);
        var bound = FormulaBinding.Bind(FormulaSyntax.Parse("=INDIRECT(\"B1\")"), book, sheet.Id, new(1, 0));
        Assert.True(FormulaSyntax.TryPrintPortable(bound, out var portable));
        Assert.Equal("=INDIRECT(\"B1\")", portable);
    }
}
