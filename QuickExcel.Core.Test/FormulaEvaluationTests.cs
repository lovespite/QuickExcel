using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class FormulaEvaluationTests
{
    [Theory]
    [InlineData("=IF(A1>0,SUM(A1:A2)*2,1/0)", "10")]
    [InlineData("=LET(x,A1+1,x*x)", "9")]
    [InlineData("=IFERROR(1/0,7)", "7")]
    [InlineData("=ROWS(A1:A2)+COLUMN()", "5")]
    public void TracesProductionEvaluatorAndLazyBranches(string formula, string value)
    {
        var sheet = Sheet("S", (0, 0, N(2)), (1, 0, N(3)), (0, 2, F(formula)));
        var trace = FormulaEvaluation.Inspect(Book(sheet), new(sheet.Id, 0, 2));
        Assert.Null(trace.Notice); Assert.Equal(value, trace.Steps[^1].Value);
        if (formula.StartsWith("=IF(")) Assert.DoesNotContain(trace.Steps, s => s.Expression == "=(1/0)");
    }
    [Fact]
    public void RequiresCurrentDependenciesAndDoesNotExecuteUnknownOrClockFunctions()
    {
        var sheet = Sheet("S", (0, 0, F("=2+2", 4)), (0, 1, F("=A1*2")), (0, 2, F("=CUSTOM_FN()")), (0, 3, F("=NOW()")));
        var snapshot = Book(sheet);
        foreach (var col in new[] { 1, 2, 3 }) { var trace = FormulaEvaluation.Inspect(snapshot, new(sheet.Id, 0, col)); Assert.Empty(trace.Steps); Assert.NotNull(trace.Notice); }
        using var session = new DocumentSession(snapshot);
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        Assert.Equal("8", FormulaEvaluation.Inspect(session.Snapshot, new(sheet.Id, 0, 1)).Steps[^1].Value);
        Assert.Throws<OperationCanceledException>(() => FormulaEvaluation.Inspect(snapshot, new(sheet.Id, 0, 1), new(true)));
    }
    [Fact]
    public void ArraysMissingTargetsErrorsAndProtectedFormulasRemainExplicit()
    {
        var sheet = Sheet("S", (0, 0, F("=SEQUENCE(2)")), (0, 2, F("=Missing!A1")), (0, 3, F("=1/0")),
            (0, 4, new(CellContent.FromFormula("=7", 1))));
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet.WithData(new() { Protection = new() })], styles: [new(), new() { FormulaHidden = true }]);
        Assert.Contains("2", FormulaEvaluation.Inspect(snapshot, new(sheet.Id, 0, 0)).Steps[^1].Value);
        Assert.Equal("#REF!", FormulaEvaluation.Inspect(snapshot, new(sheet.Id, 0, 2)).Steps[^1].Value);
        Assert.Empty(FormulaEvaluation.Inspect(snapshot, new(sheet.Id, 0, 4)).Steps);
        using var session = new DocumentSession(snapshot);
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        Assert.Equal(2, FormulaEvaluation.Errors(session.Snapshot).Count);
        Assert.Single(FormulaEvaluation.Errors(session.Snapshot, limit: 1));
        Assert.Empty(FormulaEvaluation.Errors(snapshot));
    }
    [Fact]
    public void StaleSpillReferencesAreNotPresentedAsCurrentReferenceErrors()
    {
        var sheet = Sheet("S", (0, 0, F("=SEQUENCE(2)")), (0, 2, F("=SUM(A1#)")));
        var trace = FormulaEvaluation.Inspect(Book(sheet), new(sheet.Id, 0, 2));
        Assert.NotNull(trace.Notice); Assert.Empty(trace.Steps);
    }
}
