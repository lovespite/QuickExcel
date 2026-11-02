using QuickExcel.GridGeometry;

namespace QuickExcel.Core.Test;

public sealed class ViewportAxisTests
{
    [Fact]
    public void UnboundViewportHasNoVisibleOrHitCells()
    {
        var axis = new ViewportAxis();
        Assert.Empty(axis.Visible(800)); Assert.Equal(-1, axis.HitTest(40, 800));
        Assert.Null(axis.Cell(0)); Assert.Equal(-1, axis.Neighbor(0, 1));
        axis.ScrollTo(100, 800); Assert.Equal(0, axis.Offset);
        axis.Configure(0, 30, [], [], []);
        Assert.Empty(axis.Visible(800)); Assert.Equal(-1, axis.HitTest(40, 800));
    }

    [Fact]
    public void HiddenFarTargetsChooseAdjacentModelAddressesIndependentlyOfViewport()
    {
        var axis = new ViewportAxis(); axis.Configure(200_000, 30, [], [0, 199_997, 199_998, 199_999], [0, 1]);
        Assert.Equal(199_996, axis.NearestVisible(199_998));
        Assert.Equal(199_996, axis.Neighbor(199_998, -1));
        Assert.Equal(1, axis.Neighbor(0, 1));
        axis.Configure(200_000, 30, [], [150_000, 150_001], [0]);
        Assert.Equal(150_002, axis.NearestVisible(150_000));
        Assert.Equal(149_999, axis.NearestVisible(150_001, -1));
    }
    [Fact]
    public void HiddenAndFrozenCellsShareHalfOpenHitCoordinatesWithScrolling()
    {
        var axis = new ViewportAxis();
        axis.Configure(200_000, 30, [new(4, 90)], [1], [0, 1, 2]);
        axis.ScrollTo(15, 200);
        Assert.Equal(60, axis.FrozenSize);
        Assert.Equal(new[] { 0, 2, 3, 4, 5, 6 }, axis.Visible(200).Select(c => c.ModelIndex));
        Assert.Equal(0, axis.HitTest(29.99, 200));
        Assert.Equal(2, axis.HitTest(30, 200));
        Assert.Equal(3, axis.HitTest(60, 200));
        Assert.Equal(4, axis.HitTest(75, 200));
        Assert.Equal(-1, axis.HitTest(200, 200));
        Assert.Null(axis.Cell(1));
        Assert.Equal(75, axis.Cell(4)!.Value.Start);
    }

    [Fact]
    public void OversizedCellsAlignTheirLeadingEdgeAndFarScrollStaysBounded()
    {
        var axis = new ViewportAxis();
        axis.Configure(200_000, 30, [new(50, 2000)], [], []);
        axis.EnsureVisible(50, 600);
        Assert.Equal(1500, axis.Offset);
        Assert.Single(axis.Visible(600));
        Assert.Equal(50, axis.HitTest(599, 600));
        axis.EnsureVisible(199_999, 600);
        Assert.Equal(199_999, axis.Visible(600)[^1].ModelIndex);
        Assert.InRange(axis.Visible(600).Count, 1, 21);
    }

    [Fact]
    public void ReconfigurationClearsDimensionsAndAllHiddenIsEmpty()
    {
        var axis = new ViewportAxis();
        axis.Configure(10, 30, [new(1, 200)], [], []);
        axis.Configure(10, 30, [], [], []);
        Assert.Equal(30, axis.Cell(1)!.Value.Size);
        axis.Configure(10, 30, [], Enumerable.Range(0, 10).ToHashSet(), []);
        Assert.Empty(axis.Visible(600)); Assert.Null(axis.Cell(0));
        Assert.Equal(-1, axis.HitTest(5, 600)); Assert.Equal(-1, axis.Neighbor(0, 1));
    }

    [Fact]
    public void MergedIntervalWithHiddenEndpointsUsesTheSameLogicalExtentInBothPanes()
    {
        var axis = new ViewportAxis(); axis.Configure(100, 30, [], [0, 4], [0, 1, 2]);
        axis.ScrollTo(15, 200);
        var fragments = axis.Project(0, 4, 200);
        Assert.Equal(2, fragments.Count);
        Assert.Equal(new AxisFragment(0, 90, 0, 60, true), fragments[0]);
        Assert.Equal(new AxisFragment(-15, 90, 60, 140, false), fragments[1]);
        Assert.Empty(axis.Project(4, 4, 200));
    }
}
