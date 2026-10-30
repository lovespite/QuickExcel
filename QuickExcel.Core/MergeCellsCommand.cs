using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class MergeCellsCommand(SheetId sheetId, CellRectangle range, bool unmerge = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Top > range.Bottom || range.Left > range.Right) throw new ArgumentException("选区无效。");
        var sheet = snapshot.GetSheet(sheetId);
        var merges = sheet.MergedCells.Select(CellRectangle.Parse).ToList();
        bool Overlaps(CellRectangle m) => m.Top <= range.Bottom && m.Bottom >= range.Top && m.Left <= range.Right && m.Right >= range.Left;
        if (unmerge) merges.RemoveAll(m => Overlaps(m));
        else
        {
            if (range.Top == range.Bottom && range.Left == range.Right) throw new ArgumentException("请选择多个单元格。");
            if (merges.Any(Overlaps)) throw new ArgumentException("选区与现有合并区域重叠。");
            if (sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => range.Contains(p.Row, p.Column) && (p.Row != range.Top || p.Column != range.Left)))
                throw new InvalidOperationException("合并会遮盖其他单元格的超链接或批注；操作未应用。");
            if (sheet.Cells.Any(p => range.Contains(p.Key.Row, p.Key.Column) && (p.Key.Row != range.Top || p.Key.Column != range.Left) &&
                (p.Value.Content.Kind == CellContentKind.Formula || p.Value.Value.Kind != CellValueKind.Blank)))
                throw new InvalidOperationException("合并会遮盖其他单元格内容，请先移动或清除这些内容。");
            merges.Add(range);
        }
        if (merges.Select(m => m.ToString()).SequenceEqual(sheet.MergedCells)) return new(snapshot, [], false);
        var next = new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells, sheet.Rows, sheet.Columns, merges.Select(m => m.ToString()),
            sheet.Images, sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? next : s)), [], true);
    }
}
