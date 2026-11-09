using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private void CopyFrom(WorksheetSnapshot target, CellRectangle destination, JsonElement[] args)
    {
        Arity(args, 1, 4);
        if (args.Length > 2 && args[2].ValueKind != JsonValueKind.Null && Bool(args[2]))
            throw Unsupported("copyFrom 暂不支持 skipBlanks。");
        var transpose = args.Length > 3 && args[3].ValueKind != JsonValueKind.Null && Bool(args[3]);
        var copyType = args.Length > 1 && args[1].ValueKind != JsonValueKind.Null ? EnumName("RangeCopyType", args[1]) : "all";
        var mode = copyType switch
        {
            "all" => CellPasteMode.All,
            "formulas" => CellPasteMode.Formulas,
            "values" => CellPasteMode.Values,
            "formats" => CellPasteMode.Formats,
            _ => throw Unsupported("copyFrom 暂不支持 link 复制；请使用 setHyperlink。")
        };
        var (sourceSheet, source) = SourceRange(args[0], target.Id);
        var rows = transpose ? source.Columns : source.Rows;
        var columns = transpose ? source.Rows : source.Columns;
        var single = destination.Count == 1;
        if (!single && (destination.Rows != rows || destination.Columns != columns))
            throw Unsupported("copyFrom 目前支持单格目标或与源区域等大的目标；自动平铺尚未实现。");
        var output = single
            ? new CellRectangle(destination.Top, destination.Left, destination.Top + rows - 1, destination.Left + columns - 1)
            : destination;
        WorkbookLimits.ValidatePosition(output.Top, output.Left);
        WorkbookLimits.ValidatePosition(output.Bottom, output.Right);
        Model.PrepareValueRead(sourceSheet, source, cancellation);
        var clipboard = WorkbookClipboard.Capture(Model.ValueView, sourceSheet, source);
        Write(target.Id, output, new PasteCellsCommand(clipboard, target.Id, destination, clipboard.Revision, mode, transpose));
    }

    private (SheetId Sheet, CellRectangle Range) SourceRange(JsonElement value, SheetId fallbackSheet)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            if (!value.TryGetProperty("sheetId", out var id) || !Guid.TryParse(id.GetString(), out var sheetId))
                throw Invalid("需要 Range 源。");
            var sheet = new SheetId(sheetId);
            if (value.TryGetProperty("addresses", out var addresses))
            {
                var items = addresses.EnumerateArray().Select(Text).ToArray();
                if (items.Length != 1) throw Unsupported("copyFrom 暂不支持多区域源。");
                var rect = CellRectangle.Parse(items[0]);
                return (sheet, rect);
            }
            return (sheet, Address(value, sheet));
        }
        if (value.ValueKind != JsonValueKind.String) throw Invalid("需要 Range 或区域引用字符串。");
        var text = Text(value).Trim();
        if (text.Length == 0) throw Invalid("源区域不能为空。");
        var expression = FormulaSyntax.Parse(text.StartsWith('=') ? text : "=" + text);
        if (expression is not ReferenceNode reference) throw Invalid("copyFrom 的源必须是单个矩形区域。");
        var sheetName = reference.Sheet ?? Model.View.GetSheet(fallbackSheet).Name;
        var source = Model.View.Sheets.FirstOrDefault(item => item.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ModelFailure("ItemNotFound", "找不到源区域所在的工作表。");
        var top = reference.WholeColumns ? 0 : reference.Row;
        var bottom = reference.WholeColumns ? WorkbookLimits.MaxRows - 1 : reference.IsRange ? reference.EndRow : reference.Row;
        var left = reference.Column;
        var right = reference.IsRange ? reference.EndColumn : reference.Column;
        WorkbookLimits.ValidatePosition(top, left); WorkbookLimits.ValidatePosition(bottom, right);
        if (top > bottom || left > right) throw Invalid("源区域顺序无效。");
        return (source.Id, new(top, left, bottom, right));
    }

    private object? GetSurroundingRegion(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 0); Spend(1, false);
        return RangeHandle(sheet.Id, CurrentRegion(sheet, range.Top, range.Left));
    }

    private static bool HasRecordCell(CellSnapshot? cell) =>
        cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.Value.Kind != CellValueKind.Blank);

    private static CellRectangle CurrentRegion(WorksheetSnapshot sheet, int row, int column)
    {
        if (!HasRecordCell(sheet.GetCell(row, column))) return new CellRectangle(row, column, row, column);
        var rowColumns = new Dictionary<int, HashSet<int>>();
        var columnRows = new Dictionary<int, HashSet<int>>();
        foreach (var (position, cell) in sheet.Cells)
        {
            if (!HasRecordCell(cell)) continue;
            if (!rowColumns.TryGetValue(position.Row, out var columns)) rowColumns[position.Row] = columns = [];
            columns.Add(position.Column);
            if (!columnRows.TryGetValue(position.Column, out var rows)) columnRows[position.Column] = rows = [];
            rows.Add(position.Row);
        }
        var top = row; var bottom = row; var left = column; var right = column;
        bool RowHasContent(int candidate) => rowColumns.TryGetValue(candidate, out var columns) && columns.Any(c => c >= left && c <= right);
        bool ColumnHasContent(int candidate) => columnRows.TryGetValue(candidate, out var rows) && rows.Any(r => r >= top && r <= bottom);
        bool extended;
        do
        {
            extended = false;
            while (top > 0 && RowHasContent(top - 1)) { top--; extended = true; }
            while (bottom < WorkbookLimits.MaxRows - 1 && RowHasContent(bottom + 1)) { bottom++; extended = true; }
            while (left > 0 && ColumnHasContent(left - 1)) { left--; extended = true; }
            while (right < WorkbookLimits.MaxColumns - 1 && ColumnHasContent(right + 1)) { right++; extended = true; }
        }
        while (extended);
        return new CellRectangle(top, left, bottom, right);
    }

    private object GetRangeEdge(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 1, 2);
        var direction = EnumName("KeyboardDirection", args[0]);
        var active = new CellPosition(range.Top, range.Left);
        if (args.Length == 2 && args[1].ValueKind != JsonValueKind.Null)
        {
            var cell = Address(args[1], sheet.Id);
            if (cell.Count != 1 || !range.Contains(cell.Top, cell.Left)) throw Invalid("活动格必须位于区域内。");
            active = new(cell.Top, cell.Left);
        }
        var (dr, dc) = direction switch
        {
            "up" => (-1, 0), "down" => (1, 0), "left" => (0, -1), "right" => (0, 1),
            _ => throw Invalid("方向无效。")
        };
        bool InBounds(int row, int column) => row >= 0 && row < WorkbookLimits.MaxRows && column >= 0 && column < WorkbookLimits.MaxColumns;
        bool Content(int row, int column) => HasRecordCell(sheet.GetCell(row, column));
        var currentRow = active.Row; var currentColumn = active.Column;
        var nextRow = currentRow + dr; var nextColumn = currentColumn + dc;
        if (Content(currentRow, currentColumn) && InBounds(nextRow, nextColumn) && Content(nextRow, nextColumn))
        {
            while (InBounds(nextRow + dr, nextColumn + dc) && Content(nextRow + dr, nextColumn + dc))
            { nextRow += dr; nextColumn += dc; }
        }
        else
        {
            while (InBounds(nextRow, nextColumn) && !Content(nextRow, nextColumn)) { nextRow += dr; nextColumn += dc; }
            if (!InBounds(nextRow, nextColumn))
            {
                nextRow = direction switch { "up" => 0, "down" => WorkbookLimits.MaxRows - 1, _ => active.Row };
                nextColumn = direction switch { "left" => 0, "right" => WorkbookLimits.MaxColumns - 1, _ => active.Column };
            }
        }
        Spend(1, false);
        return RangeHandle(sheet.Id, new(nextRow, nextColumn, nextRow, nextColumn));
    }

    private void GroupRange(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        Arity(args, 1);
        var option = EnumName("GroupOption", args[0]);
        var axis = option == "byRows" ? SheetAxis.Row : SheetAxis.Column;
        if (axis == SheetAxis.Row && (range.Left != 0 || range.Right != WorkbookLimits.MaxColumns - 1))
            throw Invalid("按行分组需要整行区域。");
        if (axis == SheetAxis.Column && (range.Top != 0 || range.Bottom != WorkbookLimits.MaxRows - 1))
            throw Invalid("按列分组需要整列区域。");
        var first = axis == SheetAxis.Row ? range.Top : range.Left;
        var count = axis == SheetAxis.Row ? range.Rows : range.Columns;
        if (count > MaxWriteCells) throw new ModelFailure("BudgetExceeded", "单次分组最多 20000 行列。");
        ExecuteObject(sheet.Id, new GroupAxisCommand(sheet.Id, axis, first, count, ungroup: member == "ungroup"), Math.Max(1, count));
    }

    private void GroupDetails(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        Arity(args, 1);
        var option = EnumName("GroupOption", args[0]);
        var axis = option == "byRows" ? SheetAxis.Row : SheetAxis.Column;
        var first = axis == SheetAxis.Row ? range.Top : range.Left;
        var last = axis == SheetAxis.Row ? range.Bottom : range.Right;
        var group = WorksheetOutline.Groups(sheet, axis)
            .Where(item => item.First == first && item.Last == last)
            .OrderByDescending(item => item.Level).FirstOrDefault()
            ?? throw Invalid("找不到与区域完全一致的分组。");
        ExecuteObject(sheet.Id, new SetOutlineCollapsedCommand(sheet.Id, axis, group.First, group.Last, group.Level,
            collapsed: member == "hideGroupDetails"));
    }

    private void ClearOrResetContents(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 0);
        Write(sheet.Id, range, new ClearScriptRangeCommand(sheet.Id, range, 2));
    }
}
