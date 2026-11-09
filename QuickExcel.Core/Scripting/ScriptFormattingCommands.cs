using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class ClearScriptRangeCommand(SheetId sheetId, CellRectangle range, int mode) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (range.Count > 20_000 || mode is < 0 or > 4) throw new ArgumentException("清除范围或模式无效。");
        var sheet = snapshot.GetSheet(sheetId);
        if (mode == 0 && sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => range.Contains(p.Row, p.Column)))
            throw new NotSupportedException("清除全部涉及链接或批注；请使用 contents 或 formats，附件必须显式处理。");
        var addresses = new List<CellAddress>();
        for (var r = range.Top; r <= range.Bottom; r++)
        for (var c = range.Left; c <= range.Right; c++) addresses.Add(new(sheetId, r, c));
        var current = snapshot;
        if (mode is 3 or 4)
        {
            var links = sheet.Annotations.Links.Where(p => !range.Contains(p.Key.Row, p.Key.Column)).ToImmutableDictionary();
            if (links.Count != sheet.Annotations.Links.Count)
            {
                var annotations = new WorksheetAnnotations(links, sheet.Annotations.Notes, sheet.Annotations.NoteDrawing);
                current = current.WithSheets(current.Sheets.Select(s => s.Id == sheetId ? s.WithAnnotations(annotations) : s));
            }
        }
        if (mode is 0 or 2) current = new ClearCellsCommand(addresses).Prepare(current).Snapshot;
        if (mode is 0 or 1 or 4) current = new ClearCellFormatsCommand(addresses).Prepare(current).Snapshot;
        return StructureEdits.Draft(snapshot, current) with { CalculationUnchanged = mode is 1 or 3 or 4 };
    }
}

/// <summary>Atomic heterogeneous format edits, retaining the normal rich-text/style command behavior.</summary>
public sealed class SetScriptNumberFormatsCommand(IEnumerable<(CellAddress Address, string Code)> values) : DocumentCommand
{
    private readonly (CellAddress Address, string Code)[] _values = values.Take(20_001).ToArray();
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (_values.Length > 20_000) throw new ArgumentException("数字格式修改超过预算。");
        var current = snapshot; var changes = new List<CellChange>();
        foreach (var group in _values.GroupBy(v => v.Code))
        {
            var draft = new SetCellStyleCommand(group.Select(v => v.Address), new(NumberFormatId: 0, NumberFormatCode: group.Key)).Prepare(current);
            current = draft.Snapshot; changes.AddRange(draft.Changes);
        }
        return new(current, changes.ToImmutableArray(), false, CalculationUnchanged: true);
    }
}

public sealed class SetScriptBorderCommand(SheetId sheetId, CellRectangle range, int side, int? line, string? color, int? weight) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (range.Count > 20_000 || side is < 0 or > 5) throw new ArgumentException("边框范围或方向无效。");
        var sheet = snapshot.GetSheet(sheetId); var patches = new List<(CellAddress Address, CellStylePatch Patch)>();
        for (var r = range.Top; r <= range.Bottom; r++)
        for (var c = range.Left; c <= range.Right; c++)
        {
            var s = snapshot.Styles[sheet.GetStyleId(r, c)];
            int? Apply(int old) => line ?? (weight.HasValue ? old is 3 or 4 or 6 or >= 8 ? throw new NotSupportedException("该线型不能独立保存粗细。") : weight : null);
            var top = side == 0 && r == range.Top || side == 5 && r > range.Top;
            var bottom = side == 1 && r == range.Bottom || side == 5 && r < range.Bottom;
            var left = side == 2 && c == range.Left || side == 4 && c > range.Left;
            var right = side == 3 && c == range.Right || side == 4 && c < range.Right;
            if (!top && !bottom && !left && !right) continue;
            patches.Add((new(sheetId, r, c), new(TopBorder: top ? Apply(s.TopBorder) : null, BottomBorder: bottom ? Apply(s.BottomBorder) : null,
                LeftBorder: left ? Apply(s.LeftBorder) : null, RightBorder: right ? Apply(s.RightBorder) : null,
                TopBorderColor: top ? color : null, BottomBorderColor: bottom ? color : null, LeftBorderColor: left ? color : null, RightBorderColor: right ? color : null)));
        }
        var current = snapshot; var changes = new List<CellChange>();
        foreach (var group in patches.GroupBy(p => p.Patch))
        {
            var draft = new SetCellStyleCommand(group.Select(p => p.Address), group.Key).Prepare(current);
            current = draft.Snapshot; changes.AddRange(draft.Changes);
        }
        return new(current, changes.ToImmutableArray(), false, CalculationUnchanged: true);
    }
}
