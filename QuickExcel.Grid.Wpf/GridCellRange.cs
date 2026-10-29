using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace FastWpfGrid;

public readonly record struct GridCellRange(int Top, int Left, int Bottom, int Right)
{
    public long Count => (long)(Bottom - Top + 1) * (Right - Left + 1);
    public bool Contains(int row, int column) => row >= Top && row <= Bottom && column >= Left && column <= Right;
    public bool Intersects(GridCellRange other) => Top <= other.Bottom && Bottom >= other.Top && Left <= other.Right && Right >= other.Left;
}

/// <summary>Disjoint rectangles. Selection costs depend on rectangle count, not selected cell count.</summary>
public sealed class GridRangeSelection : IEnumerable<FastGridCellAddress>
{
    private readonly List<GridCellRange> _ranges = new();
    public IReadOnlyList<GridCellRange> Ranges => _ranges.AsReadOnly();
    public long Count => _ranges.Sum(range => range.Count);
    public int RowCount => AxisCount(true);
    public int ColumnCount => AxisCount(false);
    private int AxisCount(bool rows)
    {
        int count = 0, end = -1;
        foreach (var interval in _ranges.Select(r => rows ? (Start: r.Top, End: r.Bottom) : (Start: r.Left, End: r.Right)).OrderBy(r => r.Start))
        { count += Math.Max(0, interval.End - Math.Max(end, interval.Start - 1)); end = Math.Max(end, interval.End); }
        return count;
    }
    public bool Contains(FastGridCellAddress cell) => cell.IsCell && _ranges.Any(range => range.Contains(cell.Row.GetValueOrDefault(), cell.Column.GetValueOrDefault()));
    public void Clear() => _ranges.Clear();
    public void Add(FastGridCellAddress cell) { if (cell.IsCell) Add(new GridCellRange(cell.Row.GetValueOrDefault(), cell.Column.GetValueOrDefault(), cell.Row.GetValueOrDefault(), cell.Column.GetValueOrDefault())); }
    public void Remove(FastGridCellAddress cell) { if (cell.IsCell) Remove(new GridCellRange(cell.Row.GetValueOrDefault(), cell.Column.GetValueOrDefault(), cell.Row.GetValueOrDefault(), cell.Column.GetValueOrDefault())); }
    public void Set(GridCellRange range) { Validate(range); _ranges.Clear(); _ranges.Add(range); }
    public void Set(IEnumerable<GridCellRange> ranges) { var next = ranges.ToArray(); _ranges.Clear(); foreach (var range in next) Add(range); }
    public void Add(GridCellRange range)
    {
        Validate(range);
        var pieces = new List<GridCellRange> { range };
        foreach (var existing in _ranges) pieces = pieces.SelectMany(piece => Subtract(piece, existing)).ToList();
        _ranges.AddRange(pieces);
    }
    public void Remove(GridCellRange range)
    {
        Validate(range);
        var pieces = _ranges.SelectMany(existing => Subtract(existing, range)).ToArray();
        _ranges.Clear(); _ranges.AddRange(pieces);
    }
    public void Toggle(GridCellRange range)
    {
        var added = new GridRangeSelection(); added.Set(range);
        foreach (var existing in _ranges) added.Remove(existing);
        Remove(range); _ranges.AddRange(added._ranges);
    }
    public void Transpose() => Set(_ranges.Select(r => new GridCellRange(r.Left, r.Top, r.Right, r.Bottom)).ToArray());
    private static void Validate(GridCellRange range)
    {
        if (range.Top < 0 || range.Left < 0 || range.Bottom < range.Top || range.Right < range.Left) throw new ArgumentOutOfRangeException(nameof(range));
    }
    private static IEnumerable<GridCellRange> Subtract(GridCellRange source, GridCellRange cut)
    {
        if (!source.Intersects(cut)) { yield return source; yield break; }
        var top = Math.Max(source.Top, cut.Top); var bottom = Math.Min(source.Bottom, cut.Bottom);
        var left = Math.Max(source.Left, cut.Left); var right = Math.Min(source.Right, cut.Right);
        if (source.Top < top) yield return new(source.Top, source.Left, top - 1, source.Right);
        if (bottom < source.Bottom) yield return new(bottom + 1, source.Left, source.Bottom, source.Right);
        if (source.Left < left) yield return new(top, source.Left, bottom, left - 1);
        if (right < source.Right) yield return new(top, right + 1, bottom, source.Right);
    }
    public IEnumerator<FastGridCellAddress> GetEnumerator()
    {
        if (Count > 100_000) throw new InvalidOperationException("Explicit cell enumeration exceeds 100,000 cells. Use SelectedRanges.");
        foreach (var range in _ranges)
            for (var row = range.Top; row <= range.Bottom; row++)
                for (var column = range.Left; column <= range.Right; column++) yield return new(row, column);
    }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
