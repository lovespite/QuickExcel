using QuickExcel.Formulas;

namespace QuickExcel.Domain;

/// <summary>Validated rectangular merge geometry shared by presentation and structure commands.</summary>
public readonly record struct CellRectangle(int Top, int Left, int Bottom, int Right)
{
    public bool Contains(int row, int column) => row >= Top && row <= Bottom && column >= Left && column <= Right;
    public static CellRectangle Parse(string reference)
    {
        if (FormulaSyntax.Parse("=" + reference) is not ReferenceNode node || node.Sheet is not null || node.WholeColumns)
            throw new ArgumentException("无法识别的合并区域：" + reference);
        WorkbookLimits.ValidatePosition(node.Row, node.Column); WorkbookLimits.ValidatePosition(node.EndRow, node.EndColumn);
        if (node.Row > node.EndRow || node.Column > node.EndColumn) throw new ArgumentException("合并区域顺序无效。");
        return new(node.Row, node.Column, node.EndRow, node.EndColumn);
    }
    public override string ToString() => FormulaSyntax.Print(new ReferenceNode(null, Top, Left, Bottom, Right, IsRange: true));
}
