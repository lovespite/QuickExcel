using FastWpfGrid;

namespace QuickExcel.Core.Test;

public sealed class OversizedGridCellTests
{
    [Theory]
    [InlineData(0, 4)]
    [InlineData(4, 4)]
    [InlineData(8, 4)]
    public void ScrollIntoViewKeepsAnOversizedCellVisibleFromEitherDirection(int firstVisible, int target)
    {
        var sizes = new SeriesSizes { Count = 12, DefaultSize = 100 };
        sizes.ReplaceSizeOverrides(new Dictionary<int, int> { [target] = 1768 });
        var first = sizes.ScrollInView(firstVisible, target, 800);
        Assert.Equal(target, first);
        Assert.True(sizes.IsVisible(target, first, 800));
        Assert.Equal(target + 1, sizes.ScrollInView(first, target + 1, 800));
    }

    [Fact]
    public void CellExactlyAsWideAsViewportAndOrdinaryCellsStillAlign()
    {
        var sizes = new SeriesSizes { Count = 12, DefaultSize = 100 };
        sizes.ReplaceSizeOverrides(new Dictionary<int, int> { [4] = 800 });
        Assert.Equal(4, sizes.ScrollInView(0, 4, 800));
        Assert.Equal(0, sizes.ScrollInView(0, 3, 800));
    }
}
