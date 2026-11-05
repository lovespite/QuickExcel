using System;
using FastWpfGrid;

namespace QuickExcel.GridGeometry;

/// <summary>Session-only model bounds, independent of offsets and viewport dimensions.</summary>
public sealed class GridScrollExtent
{
    public int LastRow { get; private set; }
    public int LastColumn { get; private set; }

    public void Include(int row, int column)
    {
        LastRow = Math.Max(LastRow, row);
        LastColumn = Math.Max(LastColumn, column);
    }

    public void IncludeSelection(GridCellRange range, int rows, int columns)
    {
        // Header selections and select-all do not turn the orthogonal axis into
        // a worksheet-sized scrollbar. The active cell is included separately.
        Include(range.Top == 0 && range.Bottom == rows - 1 ? 0 : range.Bottom,
            range.Left == 0 && range.Right == columns - 1 ? 0 : range.Right);
    }
}
