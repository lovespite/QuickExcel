using FastWpfGrid;

namespace QuickExcel.Core.Test;

public sealed class GridScrollRangeTests
{
    [Theory]
    [InlineData(101, 0)]
    [InlineData(100, 0)]
    [InlineData(99, 1)]
    [InlineData(50, 3)]
    [InlineData(10, 4)]
    [InlineData(0, 4)]
    public void ScrollRangeFitsTheViewportWithoutAnExtraItem(int viewport, int expected)
    {
        var sizes = new SeriesSizes { Count = 5, DefaultSize = 20 };
        Assert.Equal(expected, sizes.GetMaximumScrollIndex(viewport));
    }

    [Fact]
    public void FilteredTallRowsRemainReachableAndClearingRestoresTheRange()
    {
        var sizes = new SeriesSizes { Count = 20, DefaultSize = 10 };
        sizes.ReplaceSizeOverrides(Enumerable.Range(16, 4).ToDictionary(row => row, _ => 100));
        var original = sizes.GetMaximumScrollIndex(250);
        sizes.SetExtraordinaryIndexes(Enumerable.Range(1, 15).ToHashSet(), [0]);
        var maximum = sizes.GetMaximumScrollIndex(250);
        Assert.Equal(2, maximum);
        Assert.True(maximum < original);
        Assert.True(sizes.IsWholeInView(maximum, sizes.ScrollCount - 1, 250));
        Assert.False(sizes.IsWholeInView(maximum - 1, sizes.ScrollCount - 1, 250));
        sizes.SetExtraordinaryIndexes([], []);
        Assert.Equal(original, sizes.GetMaximumScrollIndex(250));
    }

    [Fact]
    public void EmptyFrozenOnlyAndShortFilteredSeriesHaveNoScrollRange()
    {
        var sizes = new SeriesSizes { Count = 3, DefaultSize = 20 };
        sizes.ReplaceSizeOverrides(new Dictionary<int, int> { [1] = 1000, [2] = 1000 });
        sizes.SetExtraordinaryIndexes([1, 2], []);
        Assert.Equal(0, sizes.GetMaximumScrollIndex(200));
        sizes.SetExtraordinaryIndexes([1, 2], [0]);
        Assert.Equal(0, sizes.GetMaximumScrollIndex(200));
        sizes.SetExtraordinaryIndexes([0, 1, 2], []);
        Assert.Equal(0, sizes.GetMaximumScrollIndex(200));
    }

    [Fact]
    public void OversizedLastItemCanStillBeAlignedAtTheTop()
    {
        var sizes = new SeriesSizes { Count = 3, DefaultSize = 20 };
        sizes.ReplaceSizeOverrides(new Dictionary<int, int> { [2] = 1000 });
        Assert.Equal(2, sizes.GetMaximumScrollIndex(200));
    }
}
