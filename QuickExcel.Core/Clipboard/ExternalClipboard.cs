using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Portable cell data. Styles are values, never foreign workbook IDs. Formula caches remain source results.</summary>
public sealed record ExternalClipboardCell(CellValue Value, StyleDefinition? Style = null,
    string? Formula = null, CellPosition FormulaOrigin = default, bool MissingResult = false, ExcelSerialDate? Date = null,
    string? FormulaRejection = null);

public sealed class ExternalClipboardRegion
{
    public ExternalClipboardRegion(int rows, int columns, IEnumerable<ExternalClipboardCell> cells,
        IEnumerable<CellRectangle>? merges = null, IEnumerable<string>? limitations = null, bool valuesAreDisplayText = false,
        bool hasFilterMetadata = false)
    {
        if (rows < 1 || columns < 1 || (long)rows * columns > WorkbookClipboard.MaxCells)
            throw new ArgumentException("外部粘贴需要 1–20,000 格的矩形。");
        Rows = rows; Columns = columns; ValuesAreDisplayText = valuesAreDisplayText; HasFilterMetadata = hasFilterMetadata;
        Cells = cells.Take(WorkbookClipboard.MaxCells + 1).ToImmutableArray();
        if (Cells.Length != rows * columns || Cells.Any(c => c is null)) throw new ArgumentException("剪贴板矩形不完整。");
        Merges = merges?.Take(WorkbookClipboard.MaxCells + 1).ToImmutableArray() ?? [];
        var covered = new HashSet<CellPosition>();
        foreach (var m in Merges)
        {
            if (m.Top < 0 || m.Left < 0 || m.Bottom >= rows || m.Right >= columns || m.Bottom < m.Top || m.Right < m.Left)
                throw new ArgumentException("剪贴板合并区域越界。");
            for (var r = m.Top; r <= m.Bottom; r++) for (var c = m.Left; c <= m.Right; c++)
                if (!covered.Add(new(r, c))) throw new ArgumentException("剪贴板合并区域重叠。");
        }
        Limitations = limitations?.Distinct().ToImmutableArray() ?? [];
    }
    public int Rows { get; }
    public int Columns { get; }
    public ImmutableArray<ExternalClipboardCell> Cells { get; }
    public ImmutableArray<CellRectangle> Merges { get; }
    public ImmutableArray<string> Limitations { get; }
    public bool HasFormats => Cells.All(c => c.Style is not null);
    public bool ValuesAreDisplayText { get; }
    public bool HasFilterMetadata { get; }
    public void RequireExport(bool formulas)
    {
        RequireMode(formulas ? CellPasteMode.Formulas : CellPasteMode.All);
        // SpreadsheetML/HTML merge spans omit covered cells. Never silently drop their inputs.
        foreach (var merge in Merges)
            for (var row = merge.Top; row <= merge.Bottom; row++)
                for (var column = merge.Left; column <= merge.Right; column++)
                {
                    if (row == merge.Top && column == merge.Left) continue;
                    var cell = Cells[row * Columns + column];
                    if (cell.Value.Kind != CellValueKind.Blank || cell.Formula is not null || cell.MissingResult)
                        throw new NotSupportedException("合并区域的覆盖格含内容，外部格式无法保留；请取消合并或明确选择“复制结果值”。");
                }
    }
    public void RequireMode(CellPasteMode mode)
    {
        if (!Enum.IsDefined(mode)) throw new ArgumentException("未知粘贴模式。");
        if (ValuesAreDisplayText && mode is CellPasteMode.All or CellPasteMode.Formulas)
            throw new NotSupportedException("HTML 只提供显示文本，无法恢复原始类型；请明确选择只贴值（显示文本）或只贴格式。");
        if (mode != CellPasteMode.Values && !Limitations.IsEmpty)
            throw new NotSupportedException(string.Join("；", Limitations) + "。可明确选择只贴值。");
        if (mode is CellPasteMode.All or CellPasteMode.Formats && !HasFormats)
            throw new NotSupportedException("来源未提供完整可用的格式；可明确选择只贴值。");
        if (mode is CellPasteMode.All or CellPasteMode.Values && Cells.Any(c => c.MissingResult))
            throw new InvalidOperationException("源公式缺少结果，不能粘贴为空白。");
        if (mode == CellPasteMode.Formulas && Cells.Any(c => c.FormulaRejection is not null))
            throw new NotSupportedException(Cells.First(c => c.FormulaRejection is not null).FormulaRejection + "；请明确选择只贴值。");
    }
}

/// <summary>The verified external formula subset is standard scalar functions and unqualified relative references.</summary>
public static class ExternalClipboardFormula
{
    public static FormulaNode Validate(string formula)
    {
        var expression = FormulaSyntax.Parse(formula);
        foreach (var node in FormulaSyntax.Walk(expression))
        {
            if (node is ReferenceNode r && (r.Sheet is not null || r.ColumnAbsolute || r.RowAbsolute || r.EndColumnAbsolute || r.EndRowAbsolute || r.WholeColumns) ||
                node is NameNode or StructuredReferenceNode || node is UnaryNode { Operator: "#" or "@" } ||
                node is FunctionNode f && (!StandardFunctions.Signatures.ContainsKey(f.Name) || f.Name is "INDIRECT" or "OFFSET"))
                throw new NotSupportedException("外部公式含绝对/混合引用或缺少依赖上下文的名称、工作表、函数，无法安全互粘");
        }
        if (DynamicArrayFunctions.UsesArrays(expression)) throw new NotSupportedException("外部数组公式尚未支持");
        return expression;
    }
}

/// <summary>External values never pass through CellInput and never acquire an internal clipboard identity.</summary>
public sealed class PasteExternalCellsCommand(ExternalClipboardRegion source, SheetId sheetId, CellRectangle target,
    long expectedRevision, CellPasteMode mode = CellPasteMode.All, bool allowPartial = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (snapshot.Revision != expectedRevision) throw new InvalidOperationException("读取剪贴板后文档已改变，请重试。");
        if (!allowPartial) source.RequireMode(mode);
        else if (!Enum.IsDefined(mode)) throw new ArgumentException("未知粘贴模式。");
        WorkbookClipboard.ValidateRange(target);
        var sheet = snapshot.GetSheet(sheetId);
        var scalar = source.Rows == 1 && source.Columns == 1;
        var scalarMerge = scalar && mode != CellPasteMode.Formats && sheet.MergedCells.Select(CellRectangle.Parse)
            .Any(m => m.Top == target.Top && m.Left == target.Left && (m == target || target.Top == target.Bottom && target.Left == target.Right));
        var rows = scalar ? (scalarMerge ? 1 : target.Bottom - target.Top + 1) : source.Rows;
        var columns = scalar ? (scalarMerge ? 1 : target.Right - target.Left + 1) : source.Columns;
        if (!scalar && target.Top != target.Bottom || !scalar && target.Left != target.Right)
            if (rows != target.Bottom - target.Top + 1 || columns != target.Right - target.Left + 1)
                throw new ArgumentException("选区尺寸须与粘贴数据相同，或只选择左上角单元格。");
        var output = new CellRectangle(target.Top, target.Left, checked(target.Top + rows - 1), checked(target.Left + columns - 1));
        WorkbookClipboard.ValidateRange(output);
        if (!scalarMerge) WorkbookClipboard.RequireCompleteMerges(sheet, output);
        // Partial cell paste leaves floating objects in place; it never copies or removes anchors.
        if (!allowPartial) WorkbookClipboard.RequireNoDrawings(sheet, scalarMerge ? target : output);
        WorkbookClipboard.RequireVisibleDimensions(sheet, scalarMerge ? target : output);
        if (mode is CellPasteMode.All or CellPasteMode.Formats && sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, output))))
            throw new NotSupportedException("目标涉及条件格式，完整/格式粘贴尚未支持；可选择只贴值。");
        if (mode == CellPasteMode.All && sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => output.Contains(p.Row, p.Column)))
            throw new NotSupportedException("目标含超链接/批注，完整粘贴不能安全覆盖附件；可选择只贴值。");
        var styles = snapshot.Styles.ToList(); var ids = new Dictionary<StyleDefinition, int>();
        for (var i = 0; i < styles.Count; i++) ids.TryAdd(styles[i], i);
        var edits = new List<CellEdit>();
        for (var r = 0; r < rows; r++) for (var c = 0; c < columns; c++)
        {
            var cell = source.Cells[scalar ? 0 : r * source.Columns + c]; var p = new CellPosition(output.Top + r, output.Left + c);
            // In partial mode, missing caches are holes, never manufactured blank values.
            var formula = mode == CellPasteMode.Formulas && !source.ValuesAreDisplayText && cell.FormulaRejection is null ? cell.Formula : null;
            if (formula is not null)
            {
                try { ExternalClipboardFormula.Validate(formula); formula = FormulaCopy.Translate(formula, cell.FormulaOrigin, p); }
                catch (Exception ex) when (allowPartial && ex is NotSupportedException or FormatException or ArgumentException) { formula = null; }
            }
            if (allowPartial && (mode != CellPasteMode.Formats && cell.MissingResult && formula is null || mode == CellPasteMode.Formats && cell.Style is null)) continue;
            var current = sheet.GetCell(p.Row, p.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank);
            var id = sheet.GetStyleId(p.Row, p.Column);
            if (mode is CellPasteMode.All or CellPasteMode.Formats && cell.Style is not null)
            {
                if (!ids.TryGetValue(cell.Style!, out id))
                {
                    if (styles.Count >= WorkbookLimits.MaxStyles) throw new InvalidOperationException("目标样式数量超过上限。");
                    id = styles.Count; styles.Add(cell.Style!); ids.Add(cell.Style!, id);
                }
            }
            CellContent content;
            if (mode == CellPasteMode.Formats)
            {
                if (current.RichText is not null) throw new NotSupportedException("目标富文本不能只贴格式；请先明确转换为普通文本。");
                content = current.WithStyle(id);
            }
            else if (formula is not null)
            {
                content = CellContent.FromFormula(formula, id);
            }
            else
            {
                var value = cell.Value;
                if (cell.Date is { } date)
                {
                    if (!date.TryConvertTo(snapshot.DateSystem, out var converted)) throw new NotSupportedException("日期无法转换到目标日期系统（包括虚构的 1900-02-29）。");
                    value = CellValue.FromNumber(converted.Serial);
                }
                content = CellContent.FromLiteral(value, id);
            }
            edits.Add(new(new(sheetId, p.Row, p.Column), content));
        }
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets,
            snapshot.ScriptSource, styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId,
            snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles);
        var copyMerges = mode is CellPasteMode.All or CellPasteMode.Formats && !scalarMerge &&
            !(allowPartial && mode != CellPasteMode.Formats && source.Cells.Any(c => c.MissingResult));
        var oldMerges = sheet.MergedCells.Select(CellRectangle.Parse).Where(m => ValidationSafety.Overlaps(m, output)).ToHashSet();
        var newMerges = source.Merges.Select(m => new CellRectangle(output.Top + m.Top, output.Left + m.Left, output.Top + m.Bottom, output.Left + m.Right)).ToArray();
        var mergeChanged = copyMerges && !oldMerges.SetEquals(newMerges);
        var prepared = copyMerges && oldMerges.Count > 0 ? new MergeCellsCommand(sheetId, output, unmerge: true).Prepare(styled).Snapshot : styled;
        var draft = new SetCellsCommand(edits).Prepare(prepared);
        if (copyMerges)
            foreach (var m in newMerges) draft = draft with { Snapshot = new MergeCellsCommand(sheetId, m) { PreserveCoveredContent = true }.Prepare(draft.Snapshot).Snapshot };
        if (draft.Changes.IsEmpty && !mergeChanged && !draft.StructureChanged) return new(snapshot, [], false);
        return draft with { StructureChanged = draft.StructureChanged || mergeChanged, CalculationUnchanged = mode == CellPasteMode.Formats && !mergeChanged };
    }
}
