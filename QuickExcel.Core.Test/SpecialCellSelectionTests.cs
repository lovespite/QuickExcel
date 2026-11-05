using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SpecialCellSelectionTests
{
    [Fact]
    public void EmptyTextAndFormulaEmptyTextAreNotBlankAndSpillsAreNotConstants()
    {
        var original = Sheet("S", (0, 0, T("")), (1, 0, F("=\"\"")), (2, 0, N(3)),
            (3, 0, new(CellContent.FromLiteral(CellValue.Blank))), (4, 0, F("=SEQUENCE(2)")));
        var sheet = original.WithCells(original.Cells.SetItem(new(5, 0), new(CellContent.FromLiteral(CellValue.Blank), new(CellValue.FromNumber(2), CalculationState.Fresh, 0, 0, 0), new(original.Id, 4, 0))));
        var book = Book(sheet); CellRectangle[] scope = [new(0, 0, 6, 0)];
        Assert.Equal(new[] { new CellRectangle(3, 0, 3, 0), new(6, 0, 6, 0) }, SpecialCellSelection.Find(book, sheet.Id, scope, SpecialCellKind.Blank).Ranges);
        Assert.Equal(2, SpecialCellSelection.Find(book, sheet.Id, scope, SpecialCellKind.Constant).Cells);
        Assert.Equal(2, SpecialCellSelection.Find(book, sheet.Id, scope, SpecialCellKind.Formula).Cells);
    }
    [Fact]
    public void VisibleSelectionExcludesManualHiddenColumnsAndFilteredRows()
    {
        var source = Sheet("S", (0, 0, T("Header")), (1, 0, T("A")), (2, 0, T("B")), (3, 0, T("A")));
        var sheet = new WorksheetSnapshot(source.Id, source.Name, source.Cells, rows: [new(3, new(Hidden: true))], columns: [new(1, new(Hidden: true))],
            data: new(new(new(0, 0, 3, 0), [new(0, new(values: ["A"]))])));
        var result = SpecialCellSelection.Find(Book(sheet), sheet.Id, [new(0, 0, 3, 1)], SpecialCellKind.Visible);
        Assert.Equal(2, result.Cells); Assert.Equal(new CellRectangle(0, 0, 1, 0), Assert.Single(result.Ranges));
    }
    [Fact]
    public void MergesUseOnlyAnchorAndRequireCompleteSelectableExtent()
    {
        var source = Sheet("S", (0, 0, T("merged")));
        var sheet = new WorksheetSnapshot(source.Id, source.Name, source.Cells, mergedCells: ["A1:B2"]); var book = Book(sheet);
        Assert.Equal(4, SpecialCellSelection.Find(book, sheet.Id, [new(0, 0, 1, 1)], SpecialCellKind.Constant).Cells);
        Assert.Empty(SpecialCellSelection.Find(book, sheet.Id, [new(0, 0, 1, 1)], SpecialCellKind.Blank).Ranges);
        Assert.Empty(SpecialCellSelection.Find(book, sheet.Id, [new(0, 0, 0, 0)], SpecialCellKind.Constant).Ranges);
        var protectedSheet = sheet.WithData(new() { Protection = new() { SelectLockedCells = true } });
        Assert.Empty(SpecialCellSelection.Find(Book(protectedSheet), sheet.Id, [new(0, 0, 1, 1)], SpecialCellKind.Constant).Ranges);
    }
    [Fact]
    public void SparseRemoteValuesNeverExpandIntoUsedRectangleAndOverlapsDeduplicate()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (199999, 2047, N(2))); var book = Book(sheet);
        var result = SpecialCellSelection.Find(book, sheet.Id, [new(0, 0, 199999, 2047), new(0, 0, 0, 0)], SpecialCellKind.Constant, workBudget: 20);
        Assert.Equal(2, result.Cells); Assert.Equal(2, result.Ranges.Count);
        Assert.Throws<InvalidOperationException>(() => SpecialCellSelection.Find(book, sheet.Id, [new(0, 0, 199999, 2047)], SpecialCellKind.Blank));
    }
    [Fact]
    public void CancellationBudgetAndFragmentedResultsReturnNoPartialSelection()
    {
        var sheet = Sheet("S", Enumerable.Range(0, 65).Select(i => (i * 2, 0, N(i))).ToArray()); var book = Book(sheet);
        CellRectangle[] scope = [new(0, 0, 130, 0)];
        Assert.Throws<InvalidOperationException>(() => SpecialCellSelection.Find(book, sheet.Id, scope, SpecialCellKind.Constant));
        Assert.Throws<InvalidOperationException>(() => SpecialCellSelection.Find(book, sheet.Id, scope, SpecialCellKind.Constant, workBudget: 1));
        Assert.Throws<OperationCanceledException>(() => SpecialCellSelection.Find(book, sheet.Id, scope, SpecialCellKind.Constant, new(true)));
    }
    [Fact]
    public void ErrorSelectionRejectsOldFormulaCachesAndAcceptsTypedLiteralErrors()
    {
        var stale = Sheet("S", (0, 0, F("=1/0")));
        Assert.Throws<InvalidOperationException>(() => SpecialCellSelection.Find(Book(stale), stale.Id, [new(0, 0, 0, 0)], SpecialCellKind.Error));
        var literal = Sheet("S", (0, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.DivisionByZero)))));
        Assert.Equal(1, SpecialCellSelection.Find(Book(literal), literal.Id, [new(0, 0, 0, 0)], SpecialCellKind.Error).Cells);
    }
}
