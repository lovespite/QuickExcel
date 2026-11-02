using System;
using System.Collections.Generic;
using FastWpfGrid;

namespace QuickExcel.GridGeometry;

public readonly record struct AxisCell(int ModelIndex, double Start, double Size, bool Frozen);

/// <summary>Pixel scrolling over sparse dimensions. All coordinates are unscaled DIPs.</summary>
public sealed class ViewportAxis
{
    private readonly SeriesSizes _sizes = new();
    public double Offset { get; private set; }
    public int Count => _sizes.Count;
    public double FrozenSize => _sizes.FrozenSize;
    public double ScrollSize => _sizes.GetPositionByScrollIndex(_sizes.ScrollCount);
    public void Configure(int count, int defaultSize, IEnumerable<KeyValuePair<int, int>> overrides,
        HashSet<int> hidden, HashSet<int> frozen)
    {
        if (count < 0 || defaultSize < 1) throw new ArgumentOutOfRangeException(nameof(count));
        _sizes.Clear(); _sizes.Count = count; _sizes.DefaultSize = defaultSize;
        _sizes.SetExtraordinaryIndexes(hidden, frozen);
        foreach (var item in overrides)
            if (item.Key >= 0 && item.Key < count) _sizes.PutSizeOverride(item.Key, Math.Max(1, item.Value));
        _sizes.BuildIndex();
    }
    public void ScrollTo(double value, double extent) => Offset = Math.Clamp(value, 0, Math.Max(0, ScrollSize - Math.Max(0, extent - FrozenSize)));
    public IReadOnlyList<AxisCell> Visible(double extent)
    {
        var result = new List<AxisCell>();
        for (var i = 0; i < _sizes.FrozenCount; i++)
        {
            var start = _sizes.GetPositionByRealIndex(i);
            if (start >= extent) break;
            result.Add(new(_sizes.RealToModel(i), start, _sizes.GetSizeByRealIndex(i), true));
        }
        if (extent <= FrozenSize) return result;
        var first = Math.Max(0, _sizes.GetScrollIndexOnPosition((int)Offset));
        for (var i = first; i < _sizes.ScrollCount; i++)
        {
            var start = FrozenSize + _sizes.GetPositionByScrollIndex(i) - Offset;
            if (start >= extent) break;
            result.Add(new(_sizes.RealToModel(i + _sizes.FrozenCount), start, _sizes.GetSizeByScrollIndex(i), false));
        }
        return result;
    }
    public int HitTest(double position, double extent)
    {
        if (position < 0 || position >= extent) return -1;
        // Visible intervals use half-open boundaries, including frozen boundaries.
        foreach (var cell in Visible(extent))
            if (position >= Math.Max(cell.Start, cell.Frozen ? 0 : FrozenSize) && position < cell.Start + cell.Size) return cell.ModelIndex;
        return -1;
    }
    public AxisCell? Cell(int model)
    {
        var real = _sizes.ModelToReal(model);
        if (real < 0 || real >= _sizes.RealCount) return null;
        var frozen = real < _sizes.FrozenCount;
        return new(model, _sizes.GetPositionByRealIndex(real) + (frozen ? 0 : FrozenSize - Offset), _sizes.GetSizeByRealIndex(real), frozen);
    }
    public int Neighbor(int model, int delta)
    {
        if (_sizes.RealCount == 0) return -1;
        return _sizes.RealToModel(Math.Clamp(_sizes.ModelToReal(model) + delta, 0, _sizes.RealCount - 1));
    }
    public void EnsureVisible(int model, double extent)
    {
        var cell = Cell(model);
        if (cell is not { Frozen: false } value) return;
        if (value.Size > extent - FrozenSize || value.Start < FrozenSize)
            ScrollTo(Offset + value.Start - FrozenSize, extent);
        else if (value.Start + value.Size > extent) ScrollTo(Offset + value.Start + value.Size - extent, extent);
    }
}
