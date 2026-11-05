using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class CreatePivotCommand(SheetId target, CellPosition anchor, PivotSource source, string name, int rowField,
    int? columnField, int dataField, PivotAggregation aggregation, string caption) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (snapshot.Sheets.SelectMany(s=>s.Data.Pivots).Any(p=>p.Name.Equals(name,StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException(MessageText.Get("透视表名称已存在。"));
        var cache=PivotReport.Capture(snapshot,source); var resolved=PivotReport.Resolve(snapshot,source);
        var output=PivotReport.Build(cache,anchor,rowField,columnField,dataField,aggregation,caption);
        var pivot=new WorksheetPivot(name,source with { Range=resolved.Range },output.Range,cache,rowField,columnField,dataField,aggregation,caption);
        return PivotCommands.Apply(snapshot,target,pivot,output,null);
    }
}

public sealed class RefreshPivotCommand(SheetId target, string name) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var old=snapshot.GetSheet(target).Data.Pivots.FirstOrDefault(p=>p.Name.Equals(name,StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(MessageText.Get("透视表不存在。"));
        var cache=PivotReport.Capture(snapshot,old.Source); var resolved=PivotReport.Resolve(snapshot,old.Source);
        if (!cache.Fields.SequenceEqual(old.Cache.Fields)) throw new NotSupportedException(MessageText.Get("透视来源字段已改变，不能猜测旧字段对应关系；请转换为值后重新创建透视表。"));
        var output=PivotReport.Build(cache,new(old.Range.Top,old.Range.Left),old.RowField,old.ColumnField,old.DataField,old.Aggregation,old.DataCaption,old.GrandCaption);
        var refreshed=new WorksheetPivot(old.Name,old.Source with { Range=resolved.Range },output.Range,cache,old.RowField,old.ColumnField,old.DataField,old.Aggregation,old.DataCaption,old.GrandCaption);
        return PivotCommands.Apply(snapshot,target,refreshed,output,old);
    }
}

public sealed class RemovePivotCommand(SheetId target,string name) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet=snapshot.GetSheet(target); var old=sheet.Data.Pivots.FirstOrDefault(p=>p.Name.Equals(name,StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(MessageText.Get("透视表不存在。"));
        // Removing the definition explicitly leaves the displayed results as ordinary values.
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s=>s.Id==target ? s.WithData(s.Data with { Pivots=s.Data.Pivots.Remove(old) }) : s)),[],true)
            { CalculationUnchanged=true };
    }
}

internal static class PivotCommands
{
    public static CommandDraft Apply(WorkbookSnapshot snapshot,SheetId target,WorksheetPivot pivot,PivotReport.Result output,WorksheetPivot? previous)
    {
        var sheet=snapshot.GetSheet(target); var range=output.Range;
        var source=PivotReport.Resolve(snapshot,pivot.Source);
        if (source.Sheet.Id==target && ValidationSafety.Overlaps(source.Range,range)) throw new InvalidOperationException(MessageText.Get("透视表结果不能覆盖其来源。"));
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(r=>ValidationSafety.Overlaps(r,range)) ||
            sheet.Data.Tables.Any(t=>ValidationSafety.Overlaps(t.Range,range)) ||
            sheet.Data.Pivots.Any(p=>p!=previous && ValidationSafety.Overlaps(p.Range,range)) ||
            sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p=>range.Contains(p.Row,p.Column)))
            throw new NotSupportedException(MessageText.Get("透视结果区域与合并单元格、表、其他透视表或附件重叠。"));
        foreach (var (position,cell) in sheet.Cells)
            if (range.Contains(position.Row,position.Column) && (previous is null || !previous.Range.Contains(position.Row,position.Column)) &&
                (cell.Content.Kind==CellContentKind.Formula || cell.Value.Kind!=CellValueKind.Blank || cell.SpillOwner is not null))
                throw new InvalidOperationException(MessageText.Get("透视输出会覆盖已有内容；操作未应用。"));
        var positions=output.Cells.Keys.ToHashSet();
        if (previous is not null) positions.UnionWith(sheet.Cells.Keys.Where(p=>previous.Range.Contains(p.Row,p.Column)));
        var draft=new SetCellsCommand(positions.Select(p=>new CellEdit(new(target,p.Row,p.Column),
            CellContent.FromLiteral(output.Cells.GetValueOrDefault(p),sheet.GetStyleId(p.Row,p.Column))))).Prepare(snapshot);
        var updated=draft.Snapshot.GetSheet(target);
        var pivots=previous is null ? updated.Data.Pivots.Add(pivot) : updated.Data.Pivots.Replace(previous,pivot);
        return draft with { Snapshot=draft.Snapshot.WithSheets(draft.Snapshot.Sheets.Select(s=>s.Id==target ? s.WithData(s.Data with { Pivots=pivots }) : s)), StructureChanged=true };
    }
}

internal static class PivotSafety
{
    public static void ValidateEdit(WorkbookSnapshot before,WorkbookSnapshot after,DocumentCommand command,bool structure)
    {
        if (!ReferenceEquals(before,after) && command is not SetCellStyleCommand && PivotReport.HasUnmodeledDefinitions(before))
            throw new NotSupportedException(MessageText.Get("工作簿含未建模的透视定义，无法验证结果与来源范围；保留原包，内容编辑未应用。可明确导出已有结果值副本。"));
        if (command is CreatePivotCommand or RefreshPivotCommand or RemovePivotCommand) return;
        if (!before.Sheets.Any(s=>!s.Data.Pivots.IsEmpty)) return;
        if (structure && command is not (SetSheetViewCommand or SetSheetVisibilityCommand or SetAxisMetadataCommand or SetPrintingCommand or
            SortRangeCommand or SetAutoFilterCommand or SetTableFilterCommand or ClearSortStateCommand or SetOutlineCollapsedCommand))
            throw new NotSupportedException(MessageText.Get("含透视表的工作簿暂不支持地址结构重排；可先将透视表转换为值。"));
        foreach (var sheet in before.Sheets)
        foreach (var pivot in sheet.Data.Pivots)
        {
            var changed=after.GetSheet(sheet.Id);
            foreach (var position in sheet.Cells.Keys.Concat(changed.Cells.Keys).Distinct().Where(p=>pivot.Range.Contains(p.Row,p.Column)))
                if ((sheet.GetCell(position.Row,position.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(0) !=
                    (changed.GetCell(position.Row,position.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(0))
                    throw new InvalidOperationException(MessageText.Get("不能直接修改透视结果，请修改源数据并刷新，或先转换为值。"));
        }
    }

    public static WorkbookSnapshot MarkSourceChanges(WorkbookSnapshot before,WorkbookSnapshot after)
    {
        if (!after.Sheets.Any(s=>!s.Data.Pivots.IsEmpty)) return after;
        bool Changed(WorksheetPivot pivot)
        {
            if (pivot.Stale) return false;
            var old=before.Sheets.SelectMany(s=>s.Data.Pivots).FirstOrDefault(p=>p.Name==pivot.Name);
            if (old is null || !ReferenceEquals(old.Cache,pivot.Cache)) return false; // explicit creation/refresh
            var a=PivotReport.Resolve(before,pivot.Source); var b=PivotReport.Resolve(after,pivot.Source);
            if (a.Range!=b.Range) return true;
            return a.Sheet.Cells.Keys.Concat(b.Sheet.Cells.Keys).Distinct().Where(p=>a.Range.Contains(p.Row,p.Column)).Any(p=>
            {
                var x=a.Sheet.GetCell(p.Row,p.Column); var y=b.Sheet.GetCell(p.Row,p.Column);
                return (x?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(0)!=(y?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(0) ||
                    (x?.Value ?? CellValue.Blank)!=(y?.Value ?? CellValue.Blank);
            });
        }
        return after.WithSheets(after.Sheets.Select(s=>
        {
            var items=s.Data.Pivots.Select(p=>Changed(p) ? p.WithStale(true) : p).ToImmutableArray();
            return items.SequenceEqual(s.Data.Pivots) ? s : s.WithData(s.Data with { Pivots=items });
        }));
    }
}
