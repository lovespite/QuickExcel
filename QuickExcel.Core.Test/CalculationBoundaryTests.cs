using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Formulas;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class CalculationBoundaryTests
{
    [Fact]
    public void EveryRegisteredFunctionHasIndependentExcelArityEvidence()
    {
        using var json = ScalarFormulaFixture.Load("excel-function-arity");
        var cases = json.RootElement.GetProperty("Cases").EnumerateArray().ToArray();
        Assert.Equal(StandardFunctions.Signatures.Keys.Order(), cases.Select(c => c.GetProperty("Function").GetString()!).Distinct().Order());
        foreach (var item in cases)
        {
            var name = item.GetProperty("Function").GetString()!; var count = item.GetProperty("Count").GetInt32();
            var accepted = item.GetProperty("ProductAccepted").GetBoolean();
            Assert.Equal(accepted, StandardFunctions.AcceptsArguments(name, count));
            if (accepted) continue;
            var sheet = Sheet("S", (0, 0, F(item.GetProperty("Formula").GetString()!)));
            var result = new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet)));
            Assert.Equal(CellValue.FromError(CellErrorCode.Value), C(result, sheet).Value);
        }
        Assert.False(StandardFunctions.AcceptsArguments("SUMIFS", 4));
        Assert.False(StandardFunctions.AcceptsArguments("AVERAGEIFS", 4));
        Assert.False(StandardFunctions.AcceptsArguments("COUNTIFS", 3));
    }

    [Theory]
    [InlineData(2147483648d, 1d, "")]
    [InlineData(1d, 2147483647d, "abcdef")]
    public void MidAcceptsExcelBoundaryWithZeroBasedStart(double start, double count, string expected)
    {
        var sheet = Sheet("S", (0, 0, F(FormattableString.Invariant($"=MID(\"abcdef\",{start},{count})"))));
        Assert.Equal(CellValue.FromText(expected), C(new WorkbookCalculator(new NoScripts()).Calculate(Request(Book(sheet))), sheet).Value);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void IndependentExcelScalarBoundariesMatchEditUndoAndStorage(string extension)
    {
        using var specJson = ScalarFormulaFixture.Load("calculation-boundary-spec"); using var expectedJson = ScalarFormulaFixture.Load("excel-calculation-boundary");
        var spec = specJson.RootElement; var expected = expectedJson.RootElement;
        using var session = new DocumentSession(ScalarFormulaFixture.Book(spec)); var calculator = new WorkbookCalculator(new NoScripts());
        void Check(string state) { ScalarFormulaFixture.Calculate(session, calculator); ScalarFormulaFixture.Compare(session.Snapshot, spec, expected.GetProperty(state)); }
        Check("Initial"); ScalarFormulaFixture.Edit(session, spec); Check("Edited"); Assert.True(session.Undo()); Check("Initial"); Assert.True(session.Redo()); Check("Edited");
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); using var opened = new DocumentSession(storage.Open(path));
            Assert.False(opened.Snapshot.FileState!.FormulaReadOnly); ScalarFormulaFixture.Calculate(opened, new WorkbookCalculator(new NoScripts()));
            ScalarFormulaFixture.Compare(opened.Snapshot, spec, expected.GetProperty("Edited"));
        }
        finally { File.Delete(path); }
    }
}
