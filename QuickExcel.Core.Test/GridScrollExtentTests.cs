using QuickExcel.GridGeometry;

namespace QuickExcel.Core.Test;

public sealed class GridScrollExtentTests
{
    [Fact]
    public void EmptySheetFitsOneViewportWithoutExposingTheWorksheetLimit()
    {
        var axis = new ViewportAxis(); axis.Configure(1_048_576, 31, [], [], []);
        var extent = new GridScrollExtent();
        axis.ScrollTo(double.MaxValue, 620, extent.LastRow);
        Assert.Equal(0, axis.Offset);
        Assert.Equal(20, axis.Visible(620).Count);
        Assert.Equal(1_048_575, axis.Neighbor(1_048_574, 1));
    }

    [Fact]
    public void OnlyUsedOrSelectedCellsGrowTheRangeAndReturningOrClearingDoesNotShrinkIt()
    {
        var axis = new ViewportAxis(); axis.Configure(1_048_576, 20, [], [], []);
        var extent = new GridScrollExtent(); extent.Include(99, 4);
        for (var i = 0; i < 10; i++) axis.ScrollTo(double.MaxValue, 200, extent.LastRow);
        Assert.Equal(1800, axis.Offset);
        Assert.Equal(99, extent.LastRow);
        extent.IncludeSelection(new(50, 2, 199, 8), axis.Count, 16_384);
        axis.ScrollTo(double.MaxValue, 200, extent.LastRow);
        Assert.Equal(3800, axis.Offset);
        extent.Include(0, 0); extent.IncludeSelection(new(0, 0, 0, 0), axis.Count, 16_384);
        Assert.Equal(199, extent.LastRow); Assert.Equal(8, extent.LastColumn);
        Assert.Equal(3800, axis.MaximumOffset(200, extent.LastRow));
    }

    [Fact]
    public void FullAxisSelectionsOnlyGrowTheirExplicitAxis()
    {
        var extent = new GridScrollExtent();
        extent.IncludeSelection(new(0, 0, 1_048_575, 16_383), 1_048_576, 16_384);
        Assert.Equal(0, extent.LastRow); Assert.Equal(0, extent.LastColumn);
        extent.IncludeSelection(new(60, 0, 80, 16_383), 1_048_576, 16_384);
        Assert.Equal(80, extent.LastRow); Assert.Equal(0, extent.LastColumn);
        extent.IncludeSelection(new(0, 10, 1_048_575, 15), 1_048_576, 16_384);
        Assert.Equal(80, extent.LastRow); Assert.Equal(15, extent.LastColumn);
    }

    [Fact]
    public void HiddenEndpointsFrozenPanesAndVariableSizesUseVisiblePixelBounds()
    {
        var axis = new ViewportAxis(); axis.Configure(1000, 20, [new(8, 100)], [1, 9], [0, 1, 2]);
        Assert.Equal(40, axis.FrozenSize);
        Assert.Equal(40, axis.MaximumOffset(200, 9));
        axis.ScrollTo(10000, 200, 9);
        Assert.Equal(40, axis.Offset); Assert.Equal(8, axis.Visible(200)[^1].ModelIndex);
        Assert.Equal(0, axis.MaximumOffset(200, 2));
        axis.Configure(1000, 20, [new(8, 100)], [], [0, 1, 2]);
        Assert.Equal(80, axis.MaximumOffset(200, 9));
        Assert.Equal(180, axis.MaximumOffset(100, 9));
        Assert.Equal(0, axis.MaximumOffset(400, 9));
    }

    [Fact]
    public void OversizedBoundaryCellRemainsFullyScrollable()
    {
        var axis = new ViewportAxis(); axis.Configure(1000, 20, [new(5, 2000)], [], []);
        axis.EnsureVisible(5, 400);
        Assert.Equal(100, axis.Offset);
        axis.ScrollTo(double.MaxValue, 400, 5);
        Assert.Equal(1700, axis.Offset); Assert.Equal(5, axis.Visible(400)[^1].ModelIndex);
    }

    [Fact]
    public void HiddenOrEmptyEffectiveRangesHaveNoScrollTravel()
    {
        var axis = new ViewportAxis(); axis.Configure(1000, 20, [], [0, 1, 2], []);
        Assert.Equal(0, axis.MaximumOffset(10, 2));
        axis.Configure(0, 20, [], [], []);
        Assert.Equal(0, axis.MaximumOffset(10, 2));
    }
}
