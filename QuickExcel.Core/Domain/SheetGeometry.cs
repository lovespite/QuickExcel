using QuickExcel.Core;
using QuickExcel.Formulas;

namespace QuickExcel.Domain;

/// <summary>Validated rectangular merge geometry shared by presentation and structure commands.</summary>
public readonly record struct CellRectangle(int Top, int Left, int Bottom, int Right)
{
    public int Rows => Bottom - Top + 1;
    public int Columns => Right - Left + 1;
    public long Count => (long)Rows * Columns;
    public bool Contains(int row, int column) => row >= Top && row <= Bottom && column >= Left && column <= Right;
    public bool Contains(CellRectangle other) => Contains(other.Top, other.Left) && Contains(other.Bottom, other.Right);
    public CellRectangle Offset(int rows, int columns) => new(Top + rows, Left + columns, Bottom + rows, Right + columns);
    public static CellRectangle Parse(string reference)
    {
        if (FormulaSyntax.Parse("=" + reference) is not ReferenceNode node || node.Sheet is not null || node.WholeColumns)
            throw new ArgumentException(MessageText.Get("无法识别的合并区域：") + reference);
        WorkbookLimits.ValidatePosition(node.Row, node.Column); WorkbookLimits.ValidatePosition(node.EndRow, node.EndColumn);
        if (node.Row > node.EndRow || node.Column > node.EndColumn) throw new ArgumentException(MessageText.Get("合并区域顺序无效。"));
        return new(node.Row, node.Column, node.EndRow, node.EndColumn);
    }
    public override string ToString() => FormulaSyntax.Print(new ReferenceNode(null, Top, Left, Bottom, Right, IsRange: true));
}
