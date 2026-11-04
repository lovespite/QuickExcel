using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record OutlineGroup(SheetAxis Axis, int First, int Last, int Level, int Summary, bool Collapsed);
public static class WorksheetOutline
{
    internal static void RequireEditable(WorkbookSnapshot book)
    {
        if (book.FileState?.CompatibilityDiagnostics.Any(d => d.Feature == "worksheet/autoFilter-outline") == true)
            throw new NotSupportedException("文件含尚未对齐的分组与活动筛选组合，请在 Excel 中清除该组合后再修改分组或筛选。");
    }
    internal static ColumnMetadata Column(WorksheetSnapshot sheet, int column)
    {
        var metadata = sheet.Columns.GetValueOrDefault(column) ?? new();
        // Excel treats a col element without width as zero width. Materialize the
        // current default before adding an outline-only column or summary marker.
        return metadata.Width.HasValue ? metadata : metadata with { Width = sheet.Data.DefaultColumnWidth ?? WorkbookLimits.DefaultColumnWidth(column) };
    }
    public static bool HasGroups(WorksheetSnapshot sheet) => sheet.Rows.Values.Any(r => r.OutlineLevel > 0 || r.Collapsed) || sheet.Columns.Values.Any(c => c.OutlineLevel > 0 || c.Collapsed);
    public static IReadOnlyList<OutlineGroup> Groups(WorksheetSnapshot sheet, SheetAxis axis)
    {
        if (!Enum.IsDefined(axis)) throw new ArgumentOutOfRangeException(nameof(axis));
        var forward = axis == SheetAxis.Row ? sheet.Data.Outline?.SummaryBelow ?? true : sheet.Data.Outline?.SummaryRight ?? true;
        var metadata = axis == SheetAxis.Row ? sheet.Rows.ToDictionary(p => p.Key, p => (Level: p.Value.OutlineLevel, p.Value.Collapsed)) :
            sheet.Columns.ToDictionary(p => p.Key, p => (Level: p.Value.OutlineLevel, p.Value.Collapsed));
        var result = new List<OutlineGroup>(); var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        for (var level = 1; level <= 7; level++)
        {
            var positions = metadata.Where(p => p.Value.Level >= level).Select(p => p.Key).Order().ToArray();
            for (var i = 0; i < positions.Length;)
            {
                var first = positions[i]; var last = first; i++;
                while (i < positions.Length && positions[i] == last + 1) last = positions[i++];
                var summary = forward ? last + 1 : first - 1;
                result.Add(new(axis, first, last, level, summary, summary >= 0 && summary < limit && metadata.GetValueOrDefault(summary).Collapsed));
            }
        }
        return result;
    }
    internal static WorkbookSnapshot Replace(WorkbookSnapshot book, WorksheetSnapshot sheet, IReadOnlyDictionary<int, RowMetadata> rows, IReadOnlyDictionary<int, ColumnMetadata> columns, OutlineSettings? settings)
    {
        var next = new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells, rows, columns, sheet.MergedCells, sheet.Images, sheet.ShowGridLines,
            sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations, sheet.Data with { Outline = settings });
        return book.WithSheets(book.Sheets.Select(s => s.Id == sheet.Id ? next : s));
    }
}

public sealed class GroupAxisCommand(SheetId sheetId, SheetAxis axis, int first, int count, bool ungroup = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorksheetOutline.RequireEditable(snapshot);
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (!Enum.IsDefined(axis) || first < 0 || count < 1 || count > 20000 || first > limit - count) throw new ArgumentOutOfRangeException(nameof(count));
        var sheet = snapshot.GetSheet(sheetId); var settings = sheet.Data.Outline ?? new();
        var forward = axis == SheetAxis.Row ? settings.SummaryBelow : settings.SummaryRight;
        if (forward && first + count >= limit || !forward && first == 0) throw new InvalidOperationException("分组须保留选区外的汇总行/列。");
        var groups = WorksheetOutline.Groups(sheet, axis);
        if (groups.Any(g => g.Collapsed && g.First < first + count && g.Last >= first)) throw new InvalidOperationException("请先展开相交分组，再修改分组层级。");
        var rows = sheet.Rows.ToBuilder(); var columns = sheet.Columns.ToBuilder(); var changed = false;
        for (var i = first; i < first + count; i++)
        {
            var level = axis == SheetAxis.Row ? rows.GetValueOrDefault(i)?.OutlineLevel ?? 0 : columns.GetValueOrDefault(i)?.OutlineLevel ?? 0;
            var next = ungroup ? Math.Max(0, level - 1) : level + 1;
            if (next > 7) throw new InvalidOperationException("分组最多 7 层，操作未应用。");
            changed |= next != level;
            if (axis == SheetAxis.Row) rows[i] = (rows.GetValueOrDefault(i) ?? new()) with { OutlineLevel = next };
            else columns[i] = WorksheetOutline.Column(sheet, i) with { OutlineLevel = next };
        }
        return !changed ? new(snapshot, [], false) : new(WorksheetOutline.Replace(snapshot, sheet, rows.ToImmutable(), columns.ToImmutable(), settings), [], true, true);
    }
}

public sealed class SetOutlineCollapsedCommand(SheetId sheetId, SheetAxis axis, int first, int last, int level, bool collapsed) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorksheetOutline.RequireEditable(snapshot);
        var sheet = snapshot.GetSheet(sheetId); var groups = WorksheetOutline.Groups(sheet, axis);
        var group = groups.SingleOrDefault(g => g.First == first && g.Last == last && g.Level == level) ?? throw new InvalidOperationException("分组已变化，请重新选择。");
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (group.Summary < 0 || group.Summary >= limit) throw new NotSupportedException("边界分组没有可用的汇总行/列，暂不修改。");
        if (groups.Any(g => g.Level < level && g.First <= first && g.Last >= last && g.Collapsed)) throw new InvalidOperationException("请先展开外层分组。");
        var rows = sheet.Rows.ToBuilder(); var columns = sheet.Columns.ToBuilder();
        if (axis == SheetAxis.Row) rows[group.Summary] = (rows.GetValueOrDefault(group.Summary) ?? new()) with { Collapsed = collapsed };
        else columns[group.Summary] = WorksheetOutline.Column(sheet, group.Summary) with { Collapsed = collapsed };
        for (var i = first; i <= last; i++)
        {
            var hidden = collapsed || groups.Any(g => g.Level > level && g.Summary != group.Summary && g.Collapsed && g.First <= i && g.Last >= i);
            if (axis == SheetAxis.Row) rows[i] = (rows.GetValueOrDefault(i) ?? new()) with { Hidden = hidden };
            else columns[i] = WorksheetOutline.Column(sheet, i) with { Hidden = hidden };
        }
        return new(WorksheetOutline.Replace(snapshot, sheet, rows.ToImmutable(), columns.ToImmutable(), sheet.Data.Outline ?? new()), [], true);
    }
}
