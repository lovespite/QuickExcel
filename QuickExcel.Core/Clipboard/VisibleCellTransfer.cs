using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Explicit ordered row mapping for bounded, typed transfers. Hidden columns are never compacted.</summary>
public static class VisibleCellTransfer
{
    public static int[] Rows(WorkbookSnapshot book, SheetId id, CellRectangle range, CancellationToken cancellationToken = default)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Rows < 1 || range.Columns < 1) throw new ArgumentException(MessageText.Get("选区端点无效。"));
        var sheet = book.GetSheet(id);
        if (sheet.Columns.Any(c => c.Key >= range.Left && c.Key <= range.Right && c.Value.Hidden))
            throw new NotSupportedException(MessageText.Get("可见行交换不压缩隐藏列，请先取消隐藏列或缩小选区。"));
        var filtered = WorksheetFilterEngine.Evaluate(book, id, WorksheetFilterEngine.CanUseSourceCaches(book), cancellationToken);
        var rows = new List<int>();
        for (var row = range.Top; row <= range.Bottom; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (sheet.Rows.GetValueOrDefault(row)?.Hidden == true || filtered.Contains(row)) continue;
            rows.Add(row);
            if ((long)rows.Count * range.Columns > WorkbookClipboard.MaxCells) throw new InvalidOperationException(MessageText.Get("可见行交换最多处理 20,000 个可见单元格。"));
        }
        if (rows.Count == 0) throw new InvalidOperationException(MessageText.Get("选区没有可见数据行。"));
        return rows.ToArray();
    }

    internal static void RequireAttachments(WorkbookSnapshot book, WorksheetSnapshot sheet, CellRectangle range, int[] rows, bool source, bool formats)
    {
        // Geometry and opaque discussions remain conservative over the entire enclosing rectangle.
        WorkbookClipboard.RequirePlainGeometry(sheet, range);
        ExtendedObjectSafety.ValidateCopy(book, sheet.Id, range);
        var rowSet = rows.ToHashSet();
        bool Contains(CellPosition p) => range.Contains(p.Row, p.Column) && rowSet.Contains(p.Row);
        bool Overlaps(CellRectangle r)
        {
            if (!ValidationSafety.Overlaps(r, range)) return false;
            var index = Array.BinarySearch(rows, r.Top); if (index < 0) index = ~index;
            return index < rows.Length && rows[index] <= r.Bottom;
        }
        if (sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(Contains) ||
            sheet.Cells.Any(p => Contains(p.Key) && p.Value.Content.RichText is not null) ||
            source && (sheet.Data.Validations.Any(v => v.Ranges.Any(Overlaps)) || sheet.Data.Tables.Any(t => Overlaps(t.Range))) ||
            (source || formats) && (sheet.Data.ConditionalFormats.Any(v => v.Ranges.Any(Overlaps)) ||
                book.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase)) == true))
            throw new NotSupportedException(MessageText.Get("可见行交换涉及富文本、链接、批注或无法映射的规则，请先缩小范围或使用现有专用编辑入口；整批未应用。"));
    }

    public static ExternalClipboardRegion Capture(WorkbookSnapshot book, SheetId id, CellRectangle range, bool formulas = false, CancellationToken cancellationToken = default)
    {
        var sheet = book.GetSheet(id); var rows = Rows(book, id, range, cancellationToken);
        RequireAttachments(book, sheet, range, rows, true, true);
        var cells = new List<ExternalClipboardCell>();
        foreach (var row in rows)
        {
            for (var col = range.Left; col <= range.Right; col++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = sheet.GetCell(row, col); var value = cell?.Value ?? CellValue.Blank;
                if ((cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null) &&
                    (cell.Calculation?.Value is null || cell.Calculation.InputRevision != book.Revision || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)))
                    throw new InvalidOperationException(MessageText.Get("源公式结果缺失或已过期，请先完成重算再对外复制。"));
                string? formula = null;
                if (formulas && (cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null))
                {
                    if (cell.Content.DynamicArray || cell.SpillOwner is not null || ProtectionSafety.HidesFormula(book, sheet, row, col))
                        throw new NotSupportedException(MessageText.Get("可见行公式复制不支持数组所有权或隐藏公式；请明确复制结果值。"));
                    formula = cell.Content.FormulaText; ExternalClipboardFormula.Validate(formula);
                }
                var style = book.Styles[sheet.GetStyleId(row, col)] with { FontScheme = null, NamedStyleId = 0 };
                var date = value.Kind == CellValueKind.Number && CellDisplayFormatter.IsDateFormat(style, value.NumberValue)
                    ? new ExcelSerialDate(value.NumberValue, book.DateSystem) : (ExcelSerialDate?)null;
                cells.Add(new(value, style, formula, new(row, col), Date: date));
            }
        }
        return new(rows.Length, range.Columns, cells);
    }
}

/// <summary>Validate the complete map and prepare one document transaction; failed validation never publishes any row.</summary>
public sealed class PasteVisibleCellsCommand(ExternalClipboardRegion source, SheetId sheetId, CellRectangle target,
    long expectedRevision, CellPasteMode mode = CellPasteMode.All) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (snapshot.Revision != expectedRevision) throw new InvalidOperationException(MessageText.Get("读取剪贴板后文档已改变，请重试。"));
        if (mode is not (CellPasteMode.All or CellPasteMode.Values or CellPasteMode.Formulas)) throw new ArgumentException(MessageText.Get("未知粘贴模式。"));
        source.RequireMode(mode);
        if (source.ValuesAreDisplayText || !source.Limitations.IsEmpty || !source.Merges.IsEmpty)
            throw new NotSupportedException(MessageText.Get("可见行粘贴需要无合并和未映射附件的类型化矩形，请重新复制。"));
        var rows = VisibleCellTransfer.Rows(snapshot, sheetId, target); var sheet = snapshot.GetSheet(sheetId);
        if (rows.Length != source.Rows || target.Columns != source.Columns)
            throw new ArgumentException(MessageText.Get("来源行列数与目标可见行列数不一致，请调整目标选区；整批未应用。"));
        VisibleCellTransfer.RequireAttachments(snapshot, sheet, target, rows, false, mode == CellPasteMode.All);
        var styles = snapshot.Styles.ToList(); var ids = new Dictionary<StyleDefinition, int>();
        for (var i = 0; i < styles.Count; i++) ids.TryAdd(styles[i], i);
        var edits = new List<CellEdit>();
        for (var r = 0; r < rows.Length; r++)
        {
            for (var c = 0; c < source.Columns; c++)
            {
                var cell = source.Cells[r * source.Columns + c]; var position = new CellPosition(rows[r], target.Left + c);
                var styleId = sheet.GetStyleId(position.Row, position.Column);
                if (mode == CellPasteMode.All)
                {
                    var style = cell.Style! with { NamedStyleId = 0, FontScheme = null };
                    if (!ids.TryGetValue(style, out styleId))
                    {
                        if (styles.Count >= WorkbookLimits.MaxStyles) throw new InvalidOperationException(MessageText.Get("目标样式数量超过上限。"));
                        styleId = styles.Count; styles.Add(style); ids.Add(style, styleId);
                    }
                }
                CellContent content;
                if (mode == CellPasteMode.Formulas && cell.Formula is { } formula)
                {
                    ExternalClipboardFormula.Validate(formula);
                    content = CellContent.FromFormula(FormulaCopy.Translate(formula, cell.FormulaOrigin, position), styleId);
                }
                else
                {
                    if (cell.MissingResult) throw new InvalidOperationException(MessageText.Get("源公式缺少结果，不能粘贴为空白。"));
                    var value = cell.Value;
                    if (cell.Date is { } date)
                    {
                        if (!date.TryConvertTo(snapshot.DateSystem, out var converted)) throw new NotSupportedException(MessageText.Get("日期无法转换到目标日期系统（包括虚构的 1900-02-29）。"));
                        value = CellValue.FromNumber(converted.Serial);
                    }
                    content = CellContent.FromLiteral(value, styleId);
                }
                edits.Add(new(new(sheetId, position.Row, position.Column), content));
            }
        }
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets,
            snapshot.ScriptSource, styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId,
            snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles);
        var draft = new SetCellsCommand(edits).Prepare(styled);
        return draft.Changes.IsEmpty && !draft.StructureChanged ? new(snapshot, [], false) : draft;
    }
}
