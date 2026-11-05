using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed class SetScriptTableStyleCommand(SheetId sheetId, int tableId, TableStyleInfo style) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId); var t = sheet.Data.Tables.Single(t => t.Id == tableId);
        if (style.Name != t.Style.Name && style.Name is not null && !System.Text.RegularExpressions.Regex.IsMatch(style.Name, "^TableStyle(Light([1-9]|1[0-9]|2[01])|Medium([1-9]|1[0-9]|2[0-8])|Dark([1-9]|1[01]))$"))
            throw new ArgumentException("需要内置表样式名称 TableStyleLight1–21 / Medium1–28 / Dark1–11。");
        var updated = new WorksheetTable(t.Id, t.Name, t.Range, t.Columns, t.HeaderRow, t.TotalsRow, t.TotalsRowShown, style, t.Filter, t.Sort, t.InsertRow);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(s.Data with { Tables = s.Data.Tables.Replace(t, updated) }) : s)), [], false, CalculationUnchanged: true);
    }
}

public sealed class RemoveScriptTableCommand(SheetId sheetId, int tableId, bool clear) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var table = snapshot.GetSheet(sheetId).Data.Tables.Single(t => t.Id == tableId);
        if (table.Range.Count > 20_000) throw new ArgumentException("表转换超过 20000 格。");
        if (!clear && table.Filter?.Columns.Values.Any(c => c.HasCriteria) == true)
            throw new NotSupportedException("请先清除表筛选，再转换为保留格式的普通区域。");
        // Convert all modeled references before removing their table identity.
        var transformed = StructureEdits.TransformAt(snapshot, (node, context, position) =>
        {
            var reference = node as StructuredReferenceNode ?? (node is NameNode n && n.Name.Equals(table.Name, StringComparison.OrdinalIgnoreCase) ? new StructuredReferenceNode(n.Name, Sheet: n.Sheet) : null);
            return reference is not null && StructuredReferences.Resolve(snapshot, context, position, reference)?.Table.Id == tableId
                ? clear ? new ValueNode(CellValue.FromError(CellErrorCode.Reference)) : StructuredReferences.Bind(reference, snapshot, context, position) : node;
        },
            (sheet, cells) => sheet.WithCells(cells).WithData(sheet.Data with { Tables = sheet.Data.Tables.Where(t => t.Id != tableId).ToImmutableArray() }));
        if (clear)
        {
            transformed = new ClearScriptRangeCommand(sheetId, table.Range, 0).Prepare(transformed).Snapshot;
        }
        else
        {
            // Converting a table makes its effective style explicit. Preserve
            // rich-text runs and attachments while retaining banding/borders.
            var styles = transformed.Styles.ToList(); var ids = new Dictionary<StyleDefinition, int>();
            for (var i = 0; i < styles.Count; i++) ids.TryAdd(styles[i], i);
            var source = snapshot.GetSheet(sheetId); var sheet = transformed.GetSheet(sheetId); var cells = sheet.Cells.ToBuilder();
            for (var r = table.Range.Top; r <= table.Range.Bottom; r++)
            for (var c = table.Range.Left; c <= table.Range.Right; c++)
            {
                var style = TablePresentation.GetStyle(snapshot, source, r, c);
                if (!ids.TryGetValue(style, out var id))
                {
                    if (styles.Count >= WorkbookLimits.MaxStyles) throw new InvalidOperationException("表格式转换超过样式数量上限。");
                    ids[style] = id = styles.Count; styles.Add(style);
                }
                var cell = sheet.GetCell(r, c);
                cells[new(r, c)] = new((cell?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(id), cell?.Calculation);
            }
            transformed = new(transformed.DocumentId, transformed.Revision, transformed.DateSystem,
                transformed.Sheets.Select(s => s.Id == sheetId ? s.WithCells(cells.ToImmutable()) : s), transformed.ScriptSource, styles,
                transformed.ScriptModules, transformed.FileState, transformed.DefinedNames, transformed.ActiveSheetId, transformed.Window,
                transformed.TextCompatibilityVersion, transformed.PackageFiles);
        }
        return StructureEdits.Draft(snapshot, transformed);
    }
}
