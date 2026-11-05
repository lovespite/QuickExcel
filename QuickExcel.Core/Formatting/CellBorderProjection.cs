using QuickExcel.Domain;

namespace QuickExcel.Core;

public static class CellBorderProjection
{
    // Resolve only the visible rows/columns supplied by the caller. A merged edge may
    // have gaps or different strokes, including borders contributed by its neighbors.
    public static StyleDefinition ForRow(CellRectangle range, int row, Func<int, int, StyleDefinition> styleAt)
    {
        var left = styleAt(row, range.Left);
        var right = range.Left == range.Right ? left : styleAt(row, range.Right);
        return new() { LeftBorder = left.LeftBorder, LeftBorderColor = left.LeftBorderColor,
            RightBorder = right.RightBorder, RightBorderColor = right.RightBorderColor };
    }

    public static StyleDefinition ForColumn(CellRectangle range, int column, Func<int, int, StyleDefinition> styleAt)
    {
        var top = styleAt(range.Top, column);
        var bottom = range.Top == range.Bottom ? top : styleAt(range.Bottom, column);
        return new() { TopBorder = top.TopBorder, TopBorderColor = top.TopBorderColor,
            BottomBorder = bottom.BottomBorder, BottomBorderColor = bottom.BottomBorderColor };
    }

    public static StyleDefinition ForDiagonals(StyleDefinition anchor) => new()
    {
        DiagonalBorder = anchor.DiagonalBorder, DiagonalBorderColor = anchor.DiagonalBorderColor,
        DiagonalUp = anchor.DiagonalUp, DiagonalDown = anchor.DiagonalDown
    };
}
