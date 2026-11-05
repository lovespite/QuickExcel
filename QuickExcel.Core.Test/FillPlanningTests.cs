using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class FillPlanningTests
{
    private static WorksheetSnapshot Sheet(params (int Row, int Column, string Text)[] cells) =>
        new(SheetId.New(), "S", cells.Select(c => new KeyValuePair<CellPosition, CellSnapshot>(new(c.Row, c.Column), new(CellInput.Parse(c.Text)))));

    [Fact]
    public void DoubleClickPrefersLeftRunAndStopsAtGap()
    {
        var sheet = Sheet((1, 0, "a"), (2, 0, "b"), (3, 0, "c"), (5, 0, "d"), (1, 2, "a"), (2, 2, "b"), (3, 2, "c"), (4, 2, "d"));
        Assert.Equal(new GridRange(1, 1, 3, 1), FillPlanning.DownExtent(sheet, new(1, 1, 1, 1)));
    }

    [Fact]
    public void DoubleClickFallsBackRightAndDoesNotOverwriteAnyTargetColumn()
    {
        var sheet = Sheet((1, 3, "a"), (2, 3, "b"), (3, 3, "c"), (4, 3, "d"), (3, 2, "=1"));
        Assert.Equal(new GridRange(1, 1, 2, 2), FillPlanning.DownExtent(sheet, new(1, 1, 1, 2)));
        Assert.Null(FillPlanning.DownExtent(sheet, new(2, 1, 2, 2)));
        Assert.Null(FillPlanning.DownExtent(sheet, new(4, 1, 4, 2)));
    }

    [Fact]
    public void DoubleClickFirstColumnUsesFormulaPresenceEvenWithoutCachedValues()
    {
        var sheet = Sheet((0, 1, "=1"), (1, 1, "=1"), (2, 1, "=1"));
        Assert.Equal(new GridRange(0, 0, 2, 0), FillPlanning.DownExtent(sheet, new(0, 0, 0, 0)));
    }

    [Fact]
    public void DoubleClickRefusesOversizeInsteadOfSilentlyTruncating()
    {
        var sheet = Sheet(Enumerable.Range(0, WorkbookClipboard.MaxCells + 1).Select(r => (r, 0, "1")).ToArray());
        Assert.Throws<ArgumentException>(() => FillPlanning.DownExtent(sheet, new(0, 1, 0, 1)));
        Assert.Null(FillPlanning.DownExtent(sheet, new(WorkbookLimits.MaxRows - 1, 1, WorkbookLimits.MaxRows - 1, 1)));
    }
}
