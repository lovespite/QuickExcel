using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal static class EntryNavigation
{
    public static bool HasContent(CellSnapshot? cell) => cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.Content.LiteralValue.Kind != CellValueKind.Blank);
    public static CellPosition Edge(WorksheetSnapshot sheet, CellPosition current, int dr, int dc, int rows, int columns, Func<int, bool>? filtered = null)
    {
        if (Math.Abs(dr) + Math.Abs(dc) != 1) throw new ArgumentException("One navigation axis is required.");
        bool Visible(CellPosition p) => sheet.Rows.GetValueOrDefault(p.Row)?.Hidden != true && sheet.Columns.GetValueOrDefault(p.Column)?.Hidden != true && filtered?.Invoke(p.Row) != true;
        CellPosition? Next(CellPosition p)
        {
            for (int r = p.Row + dr, c = p.Column + dc; r >= 0 && r < rows && c >= 0 && c < columns; r += dr, c += dc)
            { var candidate = new CellPosition(r, c); if (Visible(candidate)) return candidate; }
            return null;
        }
        var next = Next(current); if (next is null) return current;
        var inBlock = HasContent(sheet.GetCell(current.Row, current.Column)) && HasContent(sheet.GetCell(next.Value.Row, next.Value.Column));
        var target = next.Value;
        if (inBlock)
        {
            while (Next(target) is { } candidate && HasContent(sheet.GetCell(candidate.Row, candidate.Column))) target = candidate;
            return target;
        }
        while (!HasContent(sheet.GetCell(target.Row, target.Column)) && Next(target) is { } candidate) target = candidate;
        return target;
    }

    public static GridRange? AutoSumSource(WorksheetSnapshot sheet, CellPosition target)
    {
        // Propose the nearest contiguous numeric/formula block, above before left. The user sees the draft before committing.
        bool Numeric(int r, int c) => sheet.GetCell(r, c) is { } cell && (cell.Value.Kind == CellValueKind.Number || cell.Content.Kind == CellContentKind.Formula);
        if (target.Row > 0 && Numeric(target.Row - 1, target.Column))
        {
            var first = target.Row - 1; while (first > 0 && Numeric(first - 1, target.Column)) first--;
            return new(first, target.Column, target.Row - 1, target.Column);
        }
        if (target.Column > 0 && Numeric(target.Row, target.Column - 1))
        {
            var first = target.Column - 1; while (first > 0 && Numeric(target.Row, first - 1)) first--;
            return new(target.Row, first, target.Row, target.Column - 1);
        }
        return null;
    }
}
