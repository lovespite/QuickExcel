// QuickExcel view-only merge geometry. The host supplies model ranges, never document objects.
using System.Collections.Generic;
using System.Linq;

namespace FastWpfGrid;

partial class FastGridControl
{
    private GridCellRange[] _mergedModelRanges = [];

    public void SetMergedModelRanges(IEnumerable<GridCellRange> ranges)
    {
        var next = ranges.ToArray();
        if (_mergedModelRanges.SequenceEqual(next)) return;
        _mergedModelRanges = next;
        _selectedCells.Set(_selectedCells.Ranges.Select(ExpandMergedRange).ToArray());
        SetCurrentCell(_currentCell);
        InvalidateSelection();
        OnChangeSelectedCells(false);
    }

    public GridCellRange GetMergedModelRange(int row, int column) =>
        _mergedModelRanges.FirstOrDefault(r => r.Contains(row, column), new(row, column, row, column));

    private IEnumerable<GridCellRange> RealMergedRanges()
    {
        foreach (var merge in _mergedModelRanges)
        {
            if (merge.Top >= _modelRowCount || merge.Left >= _modelColumnCount) continue;
            var bottom = System.Math.Min(merge.Bottom, _modelRowCount - 1);
            var right = System.Math.Min(merge.Right, _modelColumnCount - 1);
            var rows = _rowSizes.GetRealRanges(IsTransposed ? merge.Left : merge.Top, IsTransposed ? right : bottom);
            var columns = _columnSizes.GetRealRanges(IsTransposed ? merge.Top : merge.Left, IsTransposed ? bottom : right);
            if (rows.Count > 0 && columns.Count > 0)
                yield return new(rows.Min(r => r.First), columns.Min(c => c.First), rows.Max(r => r.Last), columns.Max(c => c.Last));
        }
    }

    private GridCellRange ExpandMergedRange(GridCellRange range) => _mergedModelRanges.Length == 0 ? range :
        GridMergedSelection.Expand(range, RealMergedRanges().ToArray());

    private GridCellRange GetMergedRealCellRange(FastGridCellAddress cell)
    {
        var single = new GridCellRange(cell.Row.Value, cell.Column.Value, cell.Row.Value, cell.Column.Value);
        if (_mergedModelRanges.Length == 0) return single;
        return RealMergedRanges().FirstOrDefault(r => r.Contains(single.Top, single.Left), single);
    }
}
