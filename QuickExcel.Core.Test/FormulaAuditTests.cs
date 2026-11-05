using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class FormulaAuditTests
{
    [Fact]
    public void NamedCrossSheetReferencesIncludeLiteralsAndSparseRangesWithoutExpansion()
    {
        var data = Sheet("Data", (0, 0, N(2)), (1, 0, N(3))); var report = Sheet("Report", (0, 0, F("=SUM(Inputs)+Data!B1")), (1, 0, F("=A1*2")));
        var book = Book(data, report).WithDefinedNames([new("Inputs", "=Data!$A$1:$A$2")]);
        var result = FormulaAudit.Inspect(book, new(report.Id, 0, 0));
        Assert.True(result.Complete); Assert.Equal(2, result.Precedents.Count); Assert.Equal(new(0, 0, 1, 0), result.Precedents[0].Range);
        Assert.Equal(new CellAddress(report.Id, 1, 0), Assert.Single(result.Dependents).Address);
        var literal = FormulaAudit.Inspect(book, new(data.Id, 0, 0)); Assert.Empty(literal.Precedents); Assert.Equal(new CellAddress(report.Id, 0, 0), Assert.Single(literal.Dependents).Address);
        var column = Sheet("Columns", (0, 1, F("=SUM(A:A)"))); var sparse = FormulaAudit.Inspect(Book(column), new(column.Id, 0, 1));
        Assert.Single(sparse.Precedents); Assert.Equal(WorkbookLimits.MaxRows, sparse.Precedents[0].Range.Rows);
    }
    [Fact]
    public void ArrayChildrenPointToTheirAnchorAndAnchorConsumersIncludeChildReferences()
    {
        var sheet = Sheet("Array", (0, 0, new(CellContent.FromFormula("=SEQUENCE(2)", dynamicArray: true))), (0, 2, F("=A2+1")));
        using var session = new DocumentSession(Book(sheet)); Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        var child = FormulaAudit.Inspect(session.Snapshot, new(sheet.Id, 1, 0)); Assert.Equal(new CellAddress(sheet.Id, 0, 0), Assert.Single(child.Precedents).Address);
        var anchor = FormulaAudit.Inspect(session.Snapshot, new(sheet.Id, 0, 0)); Assert.Contains(anchor.Dependents, d => d.Address == new CellAddress(sheet.Id, 0, 2));
    }
    [Fact]
    public void CyclesScriptsMissingTargetsAndOldCachesAreNotReportedAsFreshCompleteCalculations()
    {
        var sheet = Sheet("S", (0, 0, F("=B1")), (0, 1, F("=A1")), (1, 0, F("=CUSTOM_FN(A1)")), (2, 0, F("=Missing!A1")), (3, 0, F("=1+1", 2)));
        var book = Book(sheet); var cycle = FormulaAudit.Inspect(book, new(sheet.Id, 0, 0));
        Assert.False(cycle.Complete); Assert.NotEmpty(cycle.Notices);
        Assert.False(FormulaAudit.Inspect(book, new(sheet.Id, 2, 0)).Complete);
        var stale = FormulaAudit.Inspect(book, new(sheet.Id, 3, 0)); Assert.Contains("旧缓存", stale.State); Assert.Equal("2", stale.Value);
        Assert.Throws<OperationCanceledException>(() => FormulaAudit.Inspect(book, new(sheet.Id, 0, 0), new(true)));
    }
    [Fact]
    public void ProtectionAndDisplayLimitDoNotLeakFormulaRelationshipsOrPretendToBeComplete()
    {
        var source = Sheet("S", (0, 0, N(1)), (1, 0, new(CellContent.FromFormula("=A1", 1))), (2, 0, F("=A1")), (3, 0, F("=A1")));
        var sheet = source.WithData(new() { Protection = new() });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { FormulaHidden = true }]);
        var protectedFormula = FormulaAudit.Inspect(book, new(sheet.Id, 1, 0)); Assert.Empty(protectedFormula.Formula); Assert.Empty(protectedFormula.Precedents);
        var consumers = FormulaAudit.Inspect(book, new(sheet.Id, 0, 0), limit: 1);
        Assert.Single(consumers.Dependents); Assert.DoesNotContain(consumers.Dependents, d => d.Address.Row == 1); Assert.False(consumers.Complete);
    }
    [Fact]
    public void StructuralChangesRebuildBindingsAndCoordinateFunctionsDoNotCreateFalseValueCycles()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (0, 1, F("=A1")), (0, 2, F("=ROW()")));
        using var session = new DocumentSession(Book(sheet)); Assert.Empty(FormulaAudit.Inspect(session.Snapshot, new(sheet.Id, 0, 2)).Precedents);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 1, false));
        var result = FormulaAudit.Inspect(session.Snapshot, new(sheet.Id, 1, 1)); Assert.Equal(new CellAddress(sheet.Id, 1, 0), Assert.Single(result.Precedents).Address);
        Assert.True(session.Undo()); Assert.Equal(new CellAddress(sheet.Id, 0, 0), Assert.Single(FormulaAudit.Inspect(session.Snapshot, new(sheet.Id, 0, 1)).Precedents).Address);
    }
}
