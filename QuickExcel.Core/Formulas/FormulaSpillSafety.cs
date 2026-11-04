using QuickExcel.Domain;

namespace QuickExcel.Core;

internal static class FormulaSpillSafety
{
    public static WorkbookSnapshot DropDerivedAfterStructure(WorkbookSnapshot snapshot) => snapshot.WithSheets(snapshot.Sheets.Select(sheet =>
    {
        if (!sheet.Cells.Values.Any(c => c.SpillOwner is not null)) return sheet;
        var cells = sheet.Cells.ToBuilder();
        foreach (var (position,cell) in sheet.Cells)
            if (cell.SpillOwner is not null)
            {
                if (cell.Content.StyleId == sheet.GetInheritedStyleId(position.Row,position.Column)) cells.Remove(position);
                else cells[position] = new(cell.Content);
            }
        return sheet.WithCells(cells.ToImmutable());
    }));

    public static void ValidateEdit(WorkbookSnapshot before, WorkbookSnapshot after, bool structureChanged)
    {
        foreach (var sheet in before.Sheets)
        {
            var next = after.Sheets.FirstOrDefault(s => s.Id == sheet.Id);
            if (next is null) continue;
            foreach (var (position, cell) in sheet.Cells)
            {
                if (cell.SpillOwner is not { } owner) continue;
                var anchor = sheet.GetCell(owner.Row, owner.Column);
                var newAnchor = next.GetCell(owner.Row, owner.Column);
                if (anchor?.Content.WithStyle(0) != newAnchor?.Content.WithStyle(0)) continue;
                var target = next.GetCell(position.Row, position.Column);
                if (target?.Content.WithStyle(0) != cell.Content.WithStyle(0))
                    throw new InvalidOperationException("不能修改数组溢出区域的一部分；请修改左上角公式或整个数组。");
                if (structureChanged && target?.SpillOwner != owner)
                    throw new InvalidOperationException("此结构操作无法保持溢出区域所有权，修改未应用。");
            }
        }
    }
}
