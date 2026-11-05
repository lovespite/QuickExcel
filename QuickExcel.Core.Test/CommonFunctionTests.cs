using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class CommonFunctionTests
{
    [Theory]
    [InlineData("=TEXTBEFORE(\"aaa\",\"aa\",-1)", "a")]
    [InlineData("=TEXTAFTER(\"甲乙甲\",\"甲\",-1)", "")]
    [InlineData("=TEXTBEFORE(\"a-b\",\"-\")", "a")]
    [InlineData("=TEXTSPLIT(\"a,b\",\",\")", "a")]
    [InlineData("=TAKE(SEQUENCE(3),-1)", "3")]
    [InlineData("=DROP(SEQUENCE(3),-1)", "1")]
    [InlineData("=SORTBY(SEQUENCE(3),SEQUENCE(3),-1)", "3")]
    public void TextAndArrayFunctionsRoundTripWithSourceAndResults(string formula, string expected)
    {
        var sheet = Sheet("Data", (0, 0, F(formula))); var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsq");
        try
        {
            using var session = new DocumentSession(Book(sheet));
            Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            Assert.False(reopened.FileState?.FormulaReadOnly == true);
            Assert.Equal(formula, reopened.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
            var result = C(new WorkbookCalculator(new NoScripts()).Calculate(Request(reopened)), reopened.Sheets[0]);
            Assert.Equal(expected, CellInput.DisplayText(result.Value!.Value));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void TextSplitSupportsDelimiterReferencesPaddingAndIncrementalShrink()
    {
        var sheet = Sheet("Data", (0, 0, T("a,b;c")), (0, 1, T(",")), (1, 1, T(";")), (0, 3, F("=TEXTSPLIT(A1,B1:B2)")));
        using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
        void Calculate() => Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        Calculate(); Assert.Equal(CellValue.FromText("c"), session.Snapshot.Sheets[0].GetCell(0, 5)!.Value);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellInput.Parse("x,y"))])); Calculate();
        Assert.Equal(CellValue.FromText("y"), session.Snapshot.Sheets[0].GetCell(0, 4)!.Value); Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 5));
        Assert.True(session.Undo()); Calculate(); Assert.Equal(CellValue.FromText("c"), session.Snapshot.Sheets[0].GetCell(0, 5)!.Value);
    }
    [Fact]
    public void ProductsDistinguishTypedRangesAndExplicitNumericConversion()
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (1, 0, T("3")), (2, 0, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            (0, 1, N(4)), (1, 1, N(5)), (2, 1, N(6)), (5, 0, F("=SUMPRODUCT(A1:A3,B1:B3)")), (6, 0, F("=SUMPRODUCT(--A1:A3,B1:B3)")));
        var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet)));
        Assert.Equal(CellValue.FromNumber(8), C(result, sheet, 5).Value);
        Assert.Equal(CellValue.FromNumber(29), C(result, sheet, 6).Value);
    }

    [Fact]
    public void IncrementalEditingUndoAndModernFunctionStorageStayConsistent()
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (0, 1, F("=IFS(A1>1,SWITCH(A1,2,20,30),TRUE,0)")));
        using var session = new DocumentSession(Book(sheet)); var calculator = new WorkbookCalculator(new NoScripts());
        void Calculate() => Assert.True(session.TryCommitCalculation(calculator.Calculate(session.BeginCalculation())));
        Calculate(); Assert.Equal(CellValue.FromNumber(20), session.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellInput.Parse("1"))])); Calculate();
        Assert.Equal(CellValue.FromNumber(0), session.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
        Assert.True(session.Undo()); Calculate(); Assert.Equal(CellValue.FromNumber(20), session.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
        var encoded = ExcelFormulaNames.Encode(sheet.GetCell(0, 1)!.Content.FormulaText);
        Assert.Contains("_xlfn.IFS", encoded); Assert.Contains("_xlfn.SWITCH", encoded);
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            Assert.False(reopened.FileState?.FormulaReadOnly == true);
            Assert.Equal(CellValue.FromNumber(20), C(calculator.Calculate(Request(reopened)), reopened.Sheets[0], 0, 1).Value);
        }
        finally { File.Delete(path); }
    }
}
