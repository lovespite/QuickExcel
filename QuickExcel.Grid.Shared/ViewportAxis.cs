using System;
using System.Collections.Generic;
using FastWpfGrid;

namespace QuickExcel.GridGeometry;

public readonly record struct AxisCell(int ModelIndex, double Start, double Size, bool Frozen);
public readonly record struct AxisFragment(double Start, double Size, double ClipStart, double ClipSize, bool Frozen);

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
    public double MaximumOffset(double extent, int? lastModel = null)
    {
        var end = ScrollSize;
        if (lastModel is { } last)
        {
            if (Count == 0 || last < 0) return 0;
            var ranges = _sizes.GetRealRanges(0, Math.Min(last, Count - 1));
            if (ranges.Count == 0 || ranges[^1].Last < _sizes.FrozenCount) return 0;
            var index = ranges[^1].Last - _sizes.FrozenCount;
            end = (double)_sizes.GetPositionByScrollIndex(index) + _sizes.GetSizeByScrollIndex(index);
        }
        return Math.Max(0, end - Math.Max(0, extent - FrozenSize));
    }
    public void ScrollTo(double value, double extent, int? lastModel = null) => Offset = Math.Clamp(value, 0, MaximumOffset(extent, lastModel));
    public IReadOnlyList<AxisCell> Visible(double extent)
    {
        var result = new List<AxisCell>();
        for (var i = 0; i < _sizes.FrozenCount; i++)
        {
            var start = _sizes.GetPositionByRealIndex(i);
            if (start >= extent) break;
            result.Add(new(_sizes.RealToModel(i), start, _sizes.GetSizeByRealIndex(i), true));
        }
        if (extent <= FrozenSize || _sizes.ScrollCount == 0) return result;
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
        var real = _sizes.ModelToReal(model);
        if (real < 0) return NearestVisible(model, delta < 0 ? -1 : 1);
        return _sizes.RealToModel(Math.Clamp(real + delta, 0, _sizes.RealCount - 1));
    }
    public int NearestVisible(int model, int direction = 1)
    {
        if (_sizes.RealCount == 0) return -1;
        model = Math.Clamp(model, 0, Count - 1);
        if (_sizes.ModelToReal(model) >= 0) return model;
        int Find(bool forward)
        {
            var ranges = _sizes.GetRealRanges(forward ? model : 0, forward ? Count - 1 : model);
            var found = -1;
            foreach (var range in ranges)
            {
                var candidate = _sizes.RealToModel(forward ? range.First : range.Last);
                if (found < 0 || (forward ? candidate < found : candidate > found)) found = candidate;
            }
            return found;
        }
        var result = Find(direction >= 0);
        return result >= 0 ? result : Find(direction < 0);
    }
    public void PreviewSize(int model, int pixels)
    {
        var real = _sizes.ModelToReal(model);
        if (real >= 0) _sizes.Resize(real, Math.Max(1, pixels));
    }
    public void EnsureVisible(int model, double extent)
    {
        var cell = Cell(model);
        if (cell is not { Frozen: false } value) return;
        if (value.Size > extent - FrozenSize || value.Start < FrozenSize)
            ScrollTo(Offset + value.Start - FrozenSize, extent);
        else if (value.Start + value.Size > extent) ScrollTo(Offset + value.Start + value.Size - extent, extent);
    }

    /// <summary>Project a model interval into frozen/scrolling clips without enumerating its cells.</summary>
    public IReadOnlyList<AxisFragment> Project(int first, int last, double extent)
    {
        var ranges = _sizes.GetRealRanges(first, last);
        var result = new List<AxisFragment>();
        if (ranges.Count == 0) return result;
        double Position(int real) => _sizes.GetPositionByRealIndex(real) + (real < _sizes.FrozenCount ? 0 : FrozenSize);
        var begin = ranges[0].First; var end = ranges[ranges.Count - 1].Last;
        var origin = Position(begin); var length = Position(end) + _sizes.GetSizeByRealIndex(end) - origin;
        void Add(bool frozen)
        {
            var start = origin - (frozen ? 0 : Offset);
            var clipStart = frozen ? 0 : FrozenSize;
            var clipEnd = frozen ? Math.Min(extent, FrozenSize) : extent;
            if (clipEnd <= clipStart || start >= clipEnd || start + length <= clipStart) return;
            result.Add(new(start, length, clipStart, clipEnd - clipStart, frozen));
        }
        if (begin < _sizes.FrozenCount) Add(true);
        if (end >= _sizes.FrozenCount) Add(false);
        return result;
    }
}
