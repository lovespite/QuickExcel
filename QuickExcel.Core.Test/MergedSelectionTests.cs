using FastWpfGrid;

namespace QuickExcel.Core.Test;

public class MergedSelectionTests
{
    [Theory]
    [InlineData(0, 5)]
    [InlineData(3, 6)]
    [InlineData(6, 7)]
    public void SelectingAnyPartIncludesTheWholeMerge(int row, int column)
    {
        var merge = new GridCellRange(0, 5, 6, 7);
        Assert.Equal(merge, GridMergedSelection.Expand(new(row, column, row, column), [merge]));
    }

    [Fact]
    public void ExpansionRevisitsEarlierMergesReachedByTheNewRectangle()
    {
        GridCellRange[] merges = [new(0, 2, 1, 3), new(0, 0, 3, 1)];
        var expected = new GridCellRange(0, 0, 3, 3);
        Assert.Equal(expected, GridMergedSelection.Expand(new(3, 1, 3, 2), merges));
        Assert.Equal(expected, GridMergedSelection.Expand(new(3, 1, 3, 2), merges.Reverse().ToArray()));
        Assert.Equal(expected, GridMergedSelection.Expand(expected, merges));
    }

    [Fact]
    public void OrdinaryCellsAndLargeRangesDoNotRequireCellEnumeration()
    {
        var merge = new GridCellRange(0, 5, 6, 7);
        var outside = new GridCellRange(7, 5, 7, 7);
        Assert.Equal(outside, GridMergedSelection.Expand(outside, [merge]));
        var all = new GridCellRange(0, 0, 99_999, 255);
        Assert.Equal(all, GridMergedSelection.Expand(all, [merge]));
        Assert.Equal(outside, GridMergedSelection.Expand(outside, []));
    }
}
