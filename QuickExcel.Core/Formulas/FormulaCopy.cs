using QuickExcel.Domain;

namespace QuickExcel.Formulas;

/// <summary>Excel A1 copy semantics. Transpose rotates a reference when at least one endpoint is fully relative.</summary>
public static class FormulaCopy
{
    public static string Translate(string formula, CellPosition source, CellPosition destination, bool transpose = false)
    {
        var changed = false;
        var expression = FormulaSyntax.Map(FormulaSyntax.Parse(formula), node =>
        {
            if (node is not ReferenceNode r) return node;
            if (transpose && r.WholeColumns)
                throw new NotSupportedException("整列引用转置需要整行引用语义，暂不支持；整批操作未应用。");
            if (transpose && (r.RowAbsolute || r.ColumnAbsolute) && (!r.IsRange || r.EndRowAbsolute || r.EndColumnAbsolute)) return node;
            (int Row, int Column) Map(int row, int column, bool rowAbsolute, bool columnAbsolute) => transpose
                ? (destination.Row + column - source.Column, destination.Column + row - source.Row)
                : (row + (rowAbsolute || r.WholeColumns ? 0 : destination.Row - source.Row),
                    column + (columnAbsolute ? 0 : destination.Column - source.Column));
            var first = Map(r.Row, r.Column, r.RowAbsolute, r.ColumnAbsolute);
            var last = r.IsRange ? Map(r.EndRow, r.EndColumn, r.EndRowAbsolute, r.EndColumnAbsolute) : first;
            var outside = first.Row < 0 || last.Row < 0 || first.Column < 0 || last.Column < 0 ||
                first.Row >= WorkbookLimits.MaxRows || last.Row >= WorkbookLimits.MaxRows ||
                first.Column >= WorkbookLimits.MaxColumns || last.Column >= WorkbookLimits.MaxColumns;
            if (outside && r.IsRange)
                throw new NotSupportedException("复制后的范围引用有端点越界，当前无法保留含 #REF! 端点的范围语法；整批操作未应用。");
            FormulaNode result = outside
                ? new ValueNode(CellValue.FromError(CellErrorCode.Reference))
                : r with { Row = first.Row, Column = first.Column, EndRow = last.Row, EndColumn = last.Column,
                    RowAbsolute = transpose ? r.ColumnAbsolute : r.RowAbsolute, ColumnAbsolute = transpose ? r.RowAbsolute : r.ColumnAbsolute,
                    EndRowAbsolute = transpose ? r.EndColumnAbsolute : r.EndRowAbsolute, EndColumnAbsolute = transpose ? r.EndRowAbsolute : r.EndColumnAbsolute };
            changed |= result != node;
            return result;
        });
        return changed ? "=" + FormulaSyntax.Print(expression) : formula;
    }
}
