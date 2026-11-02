using QuickExcel.Calculation;
using QuickExcel.Desktop;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SortExpansionTests
{
    private static WorkbookServices Services => new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts()));

    [Fact]
    public void ExpandedColumnSortMovesWholeRecordsAndRecordsSortState()
    {
        var sheet = Sheet("E",
            (0, 0, T("名称")), (0, 1, T("部门")), (0, 2, T("金额")), (0, 3, T("备注")),
            (1, 0, T("甲记录")), (1, 1, T("B")), (1, 2, N(20)), (1, 3, T("x")),
            (2, 0, T("乙记录")), (2, 1, T("A")), (2, 2, N(10)), (2, 3, T("y")),
            (3, 0, T("丙记录")), (3, 1, T("A")), (3, 2, N(30)), (3, 3, T("z")));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        doc.Session.Execute(new SortRangeCommand(new SortDescriptor(sheet.Id, new(1, 1), new(3, 1), [new(1)], expandColumns: true)));

        Assert.Equal(new[] { "乙记录", "丙记录", "甲记录" }, Enumerable.Range(1, 3).Select(row => doc.GetCellInput(0, row, 0)));
        Assert.Equal(new[] { "A", "A", "B" }, Enumerable.Range(1, 3).Select(row => doc.GetCellInput(0, row, 1)));
        Assert.Equal(new[] { "10", "30", "20" }, Enumerable.Range(1, 3).Select(row => doc.GetCellInput(0, row, 2)));
        Assert.Equal(new[] { "y", "z", "x" }, Enumerable.Range(1, 3).Select(row => doc.GetCellInput(0, row, 3)));
        Assert.Equal("名称", doc.GetCellInput(0, 0, 0));
        Assert.Equal(new CellRectangle(1, 0, 3, 3), doc.Snapshot.Sheets[0].Data.Sort!.Range);
    }

    [Fact]
    public void ExpansionIgnoresHeaderOnlyColumnsAndStopsAtBlankColumns()
    {
        var sheet = Sheet("E",
            (0, 0, T("名称")), (0, 1, T("部门")), (0, 2, T("金额")), (0, 3, T("备注")),
            (1, 0, T("甲")), (1, 1, T("B")), (1, 2, N(20)),
            (2, 0, T("乙")), (2, 1, T("A")), (2, 2, N(10)));
        Assert.Equal(new CellRectangle(0, 0, 2, 2), SortRangeCommand.ExpandColumns(sheet, new CellRectangle(0, 1, 2, 1), hasHeader: true));
        Assert.Equal(new CellRectangle(1, 0, 2, 2), SortRangeCommand.ExpandColumns(sheet, new CellRectangle(1, 1, 2, 1), hasHeader: false));
    }

    [Fact]
    public void HeaderRowStaysFixedWhileDataRowsMove()
    {
        var sheet = Sheet("E",
            (0, 0, T("名称")), (0, 1, T("数量")),
            (1, 0, T("B")), (1, 1, N(2)),
            (2, 0, T("A")), (2, 1, N(1)));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        doc.Session.Execute(new SortRangeCommand(new SortDescriptor(sheet.Id, new(0, 0), new(2, 1), [new(0)], hasHeader: true, expandColumns: true)));
        Assert.Equal("名称", doc.GetCellInput(0, 0, 0));
        Assert.Equal(new[] { "A", "B" }, Enumerable.Range(1, 2).Select(row => doc.GetCellInput(0, row, 0)));
        Assert.Equal(new[] { "1", "2" }, Enumerable.Range(1, 2).Select(row => doc.GetCellInput(0, row, 1)));
    }

    [Fact]
    public void ExpansionRejectsRangesThatGrowBeyondTheCellLimit()
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < 20; row++)
            for (var column = 0; column < 1002; column++) cells[new(row, column)] = T("x");
        var sheet = new WorksheetSnapshot(SheetId.New(), "E", cells);
        Assert.Throws<InvalidOperationException>(() => SortRangeCommand.ExpandColumns(sheet, new CellRectangle(0, 0, 19, 1000), hasHeader: false));
    }

    [Fact]
    public void SortPlanningDerivesBlockAndHeaderFromOneCell()
    {
        var sheet = Sheet("P",
            (0, 0, T("名称")), (0, 1, T("数量")),
            (1, 0, T("甲")), (1, 1, N(1)),
            (2, 0, T("乙")), (2, 1, N(2)),
            (4, 0, T("其他")), (4, 1, N(9)));
        var block = SortPlanning.Resolve(sheet, null, new(2, 0));
        Assert.Equal(new CellRectangle(0, 0, 2, 1), block.Range);
        Assert.True(block.HasHeader);
        Assert.Equal(new CellRectangle(4, 0, 4, 1), SortPlanning.Resolve(sheet, null, new(4, 0)).Range);
    }

    [Fact]
    public void SortPlanningIgnoresWholeColumnSelectionAndKeepsExplicitRows()
    {
        var sheet = Sheet("P",
            (0, 0, T("名称")), (0, 1, T("数量")),
            (1, 0, T("甲")), (1, 1, N(1)),
            (2, 0, T("乙")), (2, 1, N(2)));
        Assert.Equal(new CellRectangle(0, 0, 2, 1), SortPlanning.Resolve(sheet, new CellRectangle(0, 0, WorkbookLimits.MaxRows - 1, 0), new(1, 0)).Range);
        var explicitRows = SortPlanning.Resolve(sheet, new CellRectangle(1, 0, 2, 0), new(2, 0));
        Assert.Equal(new CellRectangle(1, 0, 2, 1), explicitRows.Range);
        Assert.False(explicitRows.HasHeader);
    }

    [Fact]
    public void SingleColumnTextBlocksDoNotGuessAHeader()
    {
        var sheet = Sheet("P", (0, 0, T("甲")), (1, 0, T("乙")));
        var block = SortPlanning.Resolve(sheet, null, new(0, 0));
        Assert.Equal(new CellRectangle(0, 0, 1, 0), block.Range);
        Assert.False(block.HasHeader);
    }

    [Fact]
    public void BlankActiveCellDoesNotInventARange()
    {
        var sheet = Sheet("P", (0, 0, T("甲")), (0, 1, N(1)));
        Assert.Equal(new CellRectangle(5, 0, 5, 0), SortPlanning.Resolve(sheet, null, new(5, 0)).Range);
    }
}
