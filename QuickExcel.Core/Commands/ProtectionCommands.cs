using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class SetProtectionCommand(SheetId sheetId, WorksheetProtection? protection) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Data.Protection == protection) return new(snapshot, [], false);
        if (sheet.Data.Protection?.HasPassword == true) throw new NotSupportedException(MessageText.Get("此工作表带密码保护，请在 Excel 中取消保护后再修改保护设置。"));
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(s.Data with { Protection = protection }) : s)), [], true, CalculationUnchanged: true);
    }
}

public sealed class SetCellProtectionCommand(SheetId sheetId, CellRectangle range, bool locked, bool hidden = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (range.Count > 20_000) throw new ArgumentException(MessageText.Get("单次保护格式最多 20,000 格。"));
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Data.Protection?.Sheet == true) throw new InvalidOperationException(MessageText.Get("请先取消工作表保护再设置锁定属性。"));
        var styles = snapshot.Styles.ToBuilder(); var cells = sheet.Cells.ToBuilder(); var changes = new List<CellChange>(); var map = new Dictionary<int, int>();
        for (var r = range.Top; r <= range.Bottom; r++)
        for (var c = range.Left; c <= range.Right; c++)
        {
            var id = sheet.GetStyleId(r, c); var before = sheet.GetCell(r, c);
            if (!map.TryGetValue(id, out var next))
            {
                var style = styles[id] with { Locked = locked, FormulaHidden = hidden }; next = styles.IndexOf(style);
                if (next < 0) { next = styles.Count; styles.Add(style); } map[id] = next;
            }
            if (id == next) continue;
            var after = new CellSnapshot((before?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(next), before?.Calculation);
            cells[new(r, c)] = after; changes.Add(new(new(sheetId, r, c), before, after));
        }
        if (changes.Count == 0) return new(snapshot, [], false);
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets, snapshot.ScriptSource,
            styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles, snapshot.LocalStorage);
        return new(styled.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithCells(cells.ToImmutable()) : s)), [.. changes], false, CalculationUnchanged: true);
    }
}

public static class ProtectionSafety
{
    public static bool CanEdit(WorkbookSnapshot book, WorksheetSnapshot sheet, int row, int column) => sheet.Data.Protection?.Sheet != true || !book.Styles[sheet.GetStyleId(row, column)].Locked;
    public static bool CanSelect(WorkbookSnapshot book, WorksheetSnapshot sheet, int row, int column) => sheet.Data.Protection is not { Sheet: true } p ||
        !(book.Styles[sheet.GetStyleId(row, column)].Locked ? p.SelectLockedCells : p.SelectUnlockedCells);
    public static bool HidesFormula(WorkbookSnapshot book, WorksheetSnapshot sheet, int row, int column) => sheet.Data.Protection?.Sheet == true && book.Styles[sheet.GetStyleId(row, column)].FormulaHidden;
    internal static void Validate(WorkbookSnapshot before, WorkbookSnapshot after, DocumentCommand command)
    {
        void Reject() => throw new InvalidOperationException(MessageText.Get("工作表已保护，此操作未被允许；工作簿未修改。"));
        // Unknown protection must never become an editable, unenforced model.
        if (before.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("sheetProtection", StringComparison.OrdinalIgnoreCase) || d.Feature.Contains("protectedRanges", StringComparison.OrdinalIgnoreCase)) == true && !ReferenceEquals(before, after)) Reject();
        foreach (var old in before.Sheets.Where(s => s.Data.Protection?.Sheet == true))
        {
            var p = old.Data.Protection!; var next = after.Sheets.FirstOrDefault(s => s.Id == old.Id);
            if (next is null) continue; // Sheet deletion is governed by workbook, not worksheet protection.
            if (p != next.Data.Protection && command is not SetProtectionCommand) Reject();
            if (command is SetProtectionCommand) continue;
            if (p.Objects && !old.Data.ThreadComments.SequenceEqual(next.Data.ThreadComments)) Reject();
            if (p.Objects && (!old.Images.SequenceEqual(next.Images) || !old.Annotations.ContentEquals(next.Annotations) || old.Data.Charts.Length != next.Data.Charts.Length || old.Data.Charts.Zip(next.Data.Charts).Any(c => !c.First.ContentEquals(c.Second)))) Reject();
            var structural = command is ShiftAxisCommand shift && shift.TargetSheet == old.Id;
            if (structural)
            {
                var axis = ((ShiftAxisCommand)command).Axis; var deletion = ((ShiftAxisCommand)command).Delete;
                if (axis == SheetAxis.Row ? deletion ? p.DeleteRows : p.InsertRows : deletion ? p.DeleteColumns : p.InsertColumns) Reject();
                if (deletion)
                {
                    var shiftCommand = (ShiftAxisCommand)command;
                    // Sparse cells are insufficient: blank locked cells also prevent deletion.
                    if (axis == SheetAxis.Row && Enumerable.Range(shiftCommand.Index, shiftCommand.Count).Any(r => before.Styles[old.Rows.GetValueOrDefault(r)?.StyleId ?? 0].Locked) ||
                        axis == SheetAxis.Column && Enumerable.Range(shiftCommand.Index, shiftCommand.Count).Any(c => before.Styles[old.Columns.GetValueOrDefault(c)?.StyleId ?? 0].Locked)) Reject();
                    if (old.Cells.Keys.Any(pos => (axis == SheetAxis.Row ? pos.Row : pos.Column) >= shiftCommand.Index &&
                        (axis == SheetAxis.Row ? pos.Row : pos.Column) < shiftCommand.Index + shiftCommand.Count && before.Styles[old.GetStyleId(pos.Row, pos.Column)].Locked)) Reject();
                    if (axis == SheetAxis.Column && old.Rows.Values.Any(r => r.StyleId is { } id && before.Styles[id].Locked)) Reject();
                }
            }
            else
            {
                if (p.FormatRows && !old.Rows.OrderBy(k => k.Key).SequenceEqual(next.Rows.OrderBy(k => k.Key))) Reject();
                if (p.FormatColumns && !old.Columns.OrderBy(k => k.Key).SequenceEqual(next.Columns.OrderBy(k => k.Key))) Reject();
                if (old.DefaultRowHeight != next.DefaultRowHeight && p.FormatRows || old.Data.DefaultColumnWidth != next.Data.DefaultColumnWidth && p.FormatColumns) Reject();
                // Axis formatting never grants permission to change inherited cell protection.
                bool ChangedProtection(int a, int b) => before.Styles[a].Locked != after.Styles[b].Locked || before.Styles[a].FormulaHidden != after.Styles[b].FormulaHidden;
                if (old.Rows.Keys.Union(next.Rows.Keys).Any(r => ChangedProtection(old.Rows.GetValueOrDefault(r)?.StyleId ?? 0, next.Rows.GetValueOrDefault(r)?.StyleId ?? 0)) ||
                    old.Columns.Keys.Union(next.Columns.Keys).Any(c => ChangedProtection(old.Columns.GetValueOrDefault(c)?.StyleId ?? 0, next.Columns.GetValueOrDefault(c)?.StyleId ?? 0))) Reject();
            }
            if (!old.MergedCells.SequenceEqual(next.MergedCells)) Reject();
            if (command is SortRangeCommand sorting && sorting.Descriptor.SheetId == old.Id)
            {
                if (p.Sort) Reject();
                var range = new CellRectangle(sorting.Descriptor.First.Row, sorting.Descriptor.First.Column, sorting.Descriptor.Last.Row, sorting.Descriptor.Last.Column);
                if (sorting.Descriptor.ExpandColumns) range = SortRangeCommand.ExpandColumns(old, range, sorting.Descriptor.HasHeader);
                for (var r = range.Top + (sorting.Descriptor.HasHeader ? 1 : 0); r <= range.Bottom; r++)
                for (var c = range.Left; c <= range.Right; c++) if (before.Styles[old.GetStyleId(r, c)].Locked) Reject();
            }
            if (old.Data.Filter is { } filter && next.Data.Filter is { } newFilter && !filter.ContentEquals(newFilter))
            { if (p.AutoFilter || filter.Range != newFilter.Range) Reject(); }
            else if ((old.Data.Filter is null) != (next.Data.Filter is null)) Reject();
            bool AllowedTableFilter(WorksheetTable a, WorksheetTable b) => command is SetTableFilterCommand && !p.AutoFilter &&
                a.Filter is not null && b.Filter is not null && a.Filter.Range == b.Filter.Range && a.ContentEquals(b.WithFilter(a.Filter));
            if (old.Data.Tables.Length != next.Data.Tables.Length || old.Data.Tables.Zip(next.Data.Tables).Any(t => !t.First.ContentEquals(t.Second) && !AllowedTableFilter(t.First, t.Second)) ||
                old.Data.Validations.Length != next.Data.Validations.Length || old.Data.Validations.Zip(next.Data.Validations).Any(v => !v.First.ContentEquals(v.Second)) ||
                !old.Data.ConditionalFormats.SequenceEqual(next.Data.ConditionalFormats) && p.FormatCells || old.Data.Outline != next.Data.Outline) Reject();
            foreach (var position in old.Cells.Keys.Union(next.Cells.Keys))
            {
                var a = old.Cells.GetValueOrDefault(position); var b = next.Cells.GetValueOrDefault(position);
                if (a?.Content == b?.Content || structural) continue;
                var oldStyle = before.Styles[old.GetStyleId(position.Row, position.Column)]; var newStyle = after.Styles[next.GetStyleId(position.Row, position.Column)];
                if ((a?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(0) != (b?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(0) && oldStyle.Locked) Reject();
                if (oldStyle != newStyle && (p.FormatCells || oldStyle.Locked != newStyle.Locked || oldStyle.FormulaHidden != newStyle.FormulaHidden)) Reject();
            }
        }
    }
}
