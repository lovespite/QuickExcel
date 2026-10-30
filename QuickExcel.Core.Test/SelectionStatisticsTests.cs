using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SelectionStatisticsTests
{
    [Fact]
    public void VisibleRowPolicyMatchesIndependentExcelAggregateAndSubtotal()
    {
        using var oracle = ScalarFormulaFixture.Load("excel-selection-statistics");
        var source = Sheet("Statistics", (1, 0, N(1.5)), (2, 0, N(2.5)), (3, 0, N(100)), (4, 0, N(200)),
            (5, 0, T("00123")), (6, 0, T("")), (7, 0, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            (8, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.NotAvailable)))), (10, 0, N(3)));
        foreach (var expected in oracle.RootElement.GetProperty("Cases").EnumerateArray())
        {
            var state = expected.GetProperty("State").GetString();
            var sheet = new WorksheetSnapshot(source.Id, source.Name, source.Cells, rows: state == "all" ? [] : [new(3, new(Hidden: true))]);
            var result = SelectionStatistics.Compute(Book(sheet), sheet.Id, [new(1, 0, 10, 0)], row => state == "filtered-and-manual" && row == 4);
            Assert.Equal(expected.GetProperty("Sum").GetDouble(), result.Sum);
            Assert.Equal(expected.GetProperty("Average").GetDouble(), result.Average);
            Assert.Equal((int)expected.GetProperty("Numbers").GetDouble(), result.Numbers);
            Assert.Equal((int)expected.GetProperty("Count").GetDouble(), result.Count);
        }
    }

    [Fact]
    public void SparseOverlappingSelectionsExcludeHiddenDimensionsAndKeepTypedCounts()
    {
        var source = Sheet("S", (0, 0, N(1)), (1, 0, N(2)), (2, 0, N(100)), (3, 0, N(200)),
            (4, 0, T("00123")), (5, 0, T("")), (6, 0, new(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            (7, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.NotAvailable)))), (8, 0, F("=3", 3)), (9, 0, F("=4")), (0, 1, N(500)));
        var sheet = new WorksheetSnapshot(source.Id, source.Name, source.Cells, rows: [new(2, new(Hidden: true))], columns: [new(1, new(Hidden: true))]);
        var book = Book(sheet);
        var result = SelectionStatistics.Compute(book, sheet.Id, [new(0, 0, 99999, 1023), new(0, 0, 9, 1)], row => row == 3);
        Assert.True(result.Complete); Assert.Equal(7, result.Count); Assert.Equal(3, result.Numbers);
        Assert.Equal(6d, result.Sum); Assert.Equal(2d, result.Average); Assert.Equal(1, result.Errors);
        Assert.Equal(1, result.OldCaches); Assert.Equal(1, result.MissingResults); Assert.Equal(11, result.VisitedCells);
        Assert.Same(sheet, book.GetSheet(sheet.Id));
    }

    [Fact]
    public void BudgetFailurePublishesNoPartialStatisticsAndCancellationRemainsObservable()
    {
        var sheet = Sheet("S", (0, 0, N(1)), (1, 0, N(2)), (2, 0, N(3))); var book = Book(sheet);
        var result = SelectionStatistics.Compute(book, sheet.Id, [new(0, 0, 99999, 1023)], comparisonBudget: 2);
        Assert.False(result.Complete); Assert.Null(result.Sum); Assert.Equal(0, result.Count); Assert.NotNull(result.Limitation);
        Assert.False(SelectionStatistics.Compute(book, sheet.Id, Enumerable.Repeat(new CellRectangle(0, 0, 0, 0), 65)).Complete);
        Assert.Throws<OperationCanceledException>(() => SelectionStatistics.Compute(book, sheet.Id, [new(0, 0, 2, 0)], cancellationToken: new(true)));
    }

    [Fact]
    public void SumCompensatesCancellationAndReportsOverflowWithoutFabricatingFiniteTotals()
    {
        var sheet = Sheet("S", (0, 0, N(1e16)), (1, 0, N(1)), (2, 0, N(-1e16))); var book = Book(sheet);
        var result = SelectionStatistics.Compute(book, sheet.Id, [new(0, 0, 2, 0)]);
        Assert.Equal(1d, result.Sum);
        var large = Sheet("S", (0, 0, N(1e308)), (1, 0, N(1e308)));
        var overflow = SelectionStatistics.Compute(Book(large), large.Id, [new(0, 0, 1, 0)]);
        Assert.Null(overflow.Sum); Assert.Equal(1e308, overflow.Average);
    }
}
