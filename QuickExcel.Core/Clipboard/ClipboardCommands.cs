using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public enum CellPasteMode { All, Values, Formulas, Formats }

/// <summary>In-process typed capture, never deserialized from an external clipboard.</summary>
public sealed class WorkbookClipboard
{
    public const int MaxCells = 20_000;
    private WorkbookClipboard(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, CellRectangle range, bool cut)
    {
        DocumentId = snapshot.DocumentId; Revision = snapshot.Revision; SheetId = sheet.Id; Range = range; IsCut = cut;
        Cells = Enumerable.Range(range.Top, Rows).SelectMany(row => Enumerable.Range(range.Left, Columns).Select(column =>
        {
            var cell = sheet.GetCell(row, column) ?? new(CellContent.FromLiteral(CellValue.Blank));
            if (ProtectionSafety.HidesFormula(snapshot, sheet, row, column) && cell.Content.Kind == CellContentKind.Formula)
                return new CellSnapshot(CellContent.FromLiteral(cell.Value, sheet.GetStyleId(row, column)));
            if (cell.SpillOwner is { } owner && !range.Contains(owner.Row,owner.Column))
                return new CellSnapshot(CellContent.FromLiteral(cell.Value,sheet.GetStyleId(row,column)));
            return new CellSnapshot(cell.Content.WithStyle(sheet.GetStyleId(row, column)), cell.Calculation, cell.SpillOwner);
        })).ToImmutableArray();
        MergedCells = sheet.MergedCells.Select(CellRectangle.Parse).Where(m => ValidationSafety.Overlaps(m, range)).ToImmutableArray();
        HasAnnotations = sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => range.Contains(p.Row, p.Column));
        HasConditionalFormats = sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, range))) ||
            snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase)) == true;
        HasOldResults = Enumerable.Range(range.Top, Rows).SelectMany(row => Enumerable.Range(range.Left, Columns)
            .Select(column => sheet.GetCell(row, column))).Any(c => c is not null && (c.Content.Kind == CellContentKind.Formula || c.SpillOwner is not null) &&
                (c.Calculation?.Value is null || c.Calculation.InputRevision != snapshot.Revision || c.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)));
    }
    public Guid DocumentId { get; }
    public long Revision { get; }
    public SheetId SheetId { get; }
    public CellRectangle Range { get; }
    public int Rows => Range.Bottom - Range.Top + 1;
    public int Columns => Range.Right - Range.Left + 1;
    public bool IsCut { get; }
    public bool HasOldResults { get; }
    public bool HasAnnotations { get; }
    public bool HasConditionalFormats { get; }
    public ImmutableArray<CellSnapshot> Cells { get; }
    public ImmutableArray<CellRectangle> MergedCells { get; }

    public static WorkbookClipboard Capture(WorkbookSnapshot snapshot, SheetId sheetId, CellRectangle range, bool cut = false)
    {
        ValidateRange(range); var sheet = snapshot.GetSheet(sheetId);
        ExtendedObjectSafety.ValidateCopy(snapshot, sheetId, range);
        if (cut) MoveCellsCommand.ValidateSource(snapshot, sheet, range);
        else { RequireCompleteMerges(sheet, range); RequireNoDrawings(sheet, range); RequireVisibleDimensions(sheet, range); }
        return new(snapshot, sheet, range, cut);
    }
    internal static void ValidateRange(CellRectangle range)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Top > range.Bottom || range.Left > range.Right || (long)(range.Bottom - range.Top + 1) * (range.Right - range.Left + 1) > MaxCells)
            throw new ArgumentException(MessageText.Get("内部复制/剪切需要不超过 20,000 格的连续矩形。"));
    }
    internal static void RequirePlainGeometry(WorksheetSnapshot sheet, CellRectangle range)
    {
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(m => ValidationSafety.Overlaps(m, range)))
            throw new NotSupportedException(MessageText.Get("此操作暂不支持合并单元格；整批操作未应用。"));
        RequireNoDrawings(sheet, range);
    }
    internal static void RequireCompleteMerges(WorksheetSnapshot sheet, CellRectangle range)
    {
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(m => ValidationSafety.Overlaps(m, range) && !range.Contains(m)))
            throw new InvalidOperationException(MessageText.Get("选区必须包含完整的合并单元格；整批操作未应用。"));
    }
    internal static void RequireNoDrawings(WorksheetSnapshot sheet, CellRectangle range)
    {
        bool Intersects(int top, int left, int bottom, int right) => top <= range.Bottom && bottom >= range.Top && left <= range.Right && right >= range.Left;
        if (sheet.Images.Select(i => i.Anchor).Concat(sheet.Data.Charts.Select(c => c.Anchor)).Any(a =>
            {
                var b = a.Bounds(sheet); if (b.Width <= 0 || b.Height <= 0) return false;
                var extent = DrawingGeometry.FromBounds(sheet, b, DrawingAnchorKind.TwoCell);
                return Intersects(extent.Row, extent.Column, extent.EndRow - (extent.EndRowOffset == 0 ? 1 : 0), extent.EndColumn - (extent.EndColumnOffset == 0 ? 1 : 0));
            }))
            throw new NotSupportedException(MessageText.Get("内部复制/粘贴区域含图片或图表，尚不能完整复制其结构；整批操作未应用。"));
    }
    internal static void RequireVisibleDimensions(WorksheetSnapshot sheet, CellRectangle range)
    {
        if (sheet.Rows.Any(r => r.Key >= range.Top && r.Key <= range.Bottom && r.Value.Hidden) ||
            sheet.Columns.Any(c => c.Key >= range.Left && c.Key <= range.Right && c.Value.Hidden))
            throw new InvalidOperationException(MessageText.Get("内部复制/剪切/粘贴区域含隐藏行列。请先取消隐藏；整批操作未应用。"));
    }
}

/// <summary>One atomic paste. The caller advances its lease only after its own successful paste; any other edit expires it.</summary>
public sealed class PasteCellsCommand(WorkbookClipboard clipboard, SheetId sheetId, CellRectangle target,
    long expectedRevision, CellPasteMode mode = CellPasteMode.All, bool transpose = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(clipboard);
        if (!Enum.IsDefined(mode)) throw new ArgumentException(MessageText.Get("未知粘贴模式。"));
        if (transpose && mode is CellPasteMode.All or CellPasteMode.Formulas && clipboard.Cells.Any(c => c.Content.DynamicArray || c.SpillOwner is not null))
            throw new NotSupportedException(MessageText.Get("数组公式的转置粘贴尚未支持；可明确选择转置粘贴值。"));
        if (snapshot.DocumentId != clipboard.DocumentId || snapshot.Revision != expectedRevision || expectedRevision < clipboard.Revision)
            throw new InvalidOperationException(MessageText.Get("复制后文档已改变，请重新复制或剪切；整批操作未应用。"));
        WorkbookClipboard.ValidateRange(target);
        var sheet = snapshot.GetSheet(sheetId);
        // A scalar pasted into a selected merged cell edits its anchor, like ordinary entry.
        var scalarMerge = mode != CellPasteMode.Formats && !clipboard.IsCut && clipboard.Rows == 1 && clipboard.Columns == 1 &&
            sheet.MergedCells.Select(CellRectangle.Parse).Any(m => m.Top == target.Top && m.Left == target.Left &&
                (m == target || target.Top == target.Bottom && target.Left == target.Right));
        var destination = scalarMerge ? new CellRectangle(target.Top, target.Left, target.Top, target.Left) : target;
        var rows = transpose ? clipboard.Columns : clipboard.Rows; var columns = transpose ? clipboard.Rows : clipboard.Columns;
        var single = destination.Top == destination.Bottom && destination.Left == destination.Right;
        var fill = clipboard.Rows == 1 && clipboard.Columns == 1 && !clipboard.IsCut;
        if (fill) { rows = destination.Bottom - destination.Top + 1; columns = destination.Right - destination.Left + 1; }
        var cutToMerge = clipboard.IsCut && !single && sheet.MergedCells.Select(CellRectangle.Parse).Contains(target);
        if (!single && !fill && !cutToMerge && (destination.Rows != rows || destination.Columns != columns))
            throw new ArgumentException(MessageText.Get("选区尺寸须与粘贴数据相同，或只选择左上角单元格。"));
        var output = new CellRectangle(destination.Top, destination.Left, checked(destination.Top + rows - 1), checked(destination.Left + columns - 1));
        WorkbookClipboard.ValidateRange(output);
        if (clipboard.IsCut)
        {
            if (mode != CellPasteMode.All || transpose) throw new NotSupportedException(MessageText.Get("剪切仅支持完整粘贴，不支持选择性粘贴或转置；源区域尚未修改。"));
            if (expectedRevision != clipboard.Revision) throw new InvalidOperationException(MessageText.Get("剪切后文档已改变，请重新剪切。"));
            return new MoveCellsCommand(clipboard.SheetId, new(clipboard.Range.Top, clipboard.Range.Left),
                new(clipboard.Range.Bottom, clipboard.Range.Right), sheetId, new(output.Top, output.Left)).Prepare(snapshot);
        }
        if (!scalarMerge) WorkbookClipboard.RequireCompleteMerges(sheet, output);
        WorkbookClipboard.RequireNoDrawings(sheet, scalarMerge ? target : output);
        WorkbookClipboard.RequireVisibleDimensions(sheet, scalarMerge ? target : output);
        if (mode is CellPasteMode.All or CellPasteMode.Formats && (clipboard.HasConditionalFormats ||
            sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, output)))))
            throw new NotSupportedException(MessageText.Get("完整/格式粘贴涉及条件格式，尚不能复制或拆分其作用范围；可明确选择只贴值或公式，规则可在条件格式管理器中编辑。"));
        if (mode == CellPasteMode.All && (clipboard.HasAnnotations ||
            sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => output.Contains(p.Row, p.Column))))
            throw new NotSupportedException(MessageText.Get("完整粘贴涉及超链接/批注，尚不能安全变换附件；可明确选择只贴值、公式或格式。"));
        if (mode == CellPasteMode.Values && clipboard.Cells.Any(c => c.Content.Kind == CellContentKind.Formula && c.Calculation?.Value is null))
            throw new InvalidOperationException(MessageText.Get("源公式缺少结果，不能把它粘贴为空白；请先完成重算。"));
        var edits = new List<CellEdit>();
        for (var row = 0; row < rows; row++)
        for (var column = 0; column < columns; column++)
        {
            var sr = fill ? 0 : transpose ? column : row; var sc = fill ? 0 : transpose ? row : column;
            var original = clipboard.Cells[sr * clipboard.Columns + sc];
            var position = new CellPosition(output.Top + row, output.Left + column);
            var current = sheet.GetCell(position.Row, position.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank);
            var style = mode is CellPasteMode.All or CellPasteMode.Formats ? original.Content.StyleId : sheet.GetStyleId(position.Row, position.Column);
            CellContent content;
            if (mode == CellPasteMode.Formats)
            {
                if (original.Content.RichText is not null || current.RichText is not null)
                    throw new NotSupportedException(MessageText.Get("富文本分段格式不能通过只贴格式准确映射；请先明确转换为普通文本，整批操作未应用。"));
                content = current.WithStyle(style);
            }
            else if (mode == CellPasteMode.Values) content = CellContent.FromLiteral(original.Value, style);
            else if (original.Content.Kind == CellContentKind.Formula)
            {
                content = CellContent.FromFormula(CopyFormula(snapshot, clipboard.SheetId, sheetId, original.Content.FormulaText,
                    new(clipboard.Range.Top + sr, clipboard.Range.Left + sc), position, transpose), style, original.Content.DynamicArray);
            }
            else content = mode == CellPasteMode.All || original.SpillOwner is not null ? original.Content.WithStyle(style) : CellContent.FromLiteral(original.Value, style);
            if (mode == CellPasteMode.Formats && content == sheet.GetCell(position.Row, position.Column)?.Content) continue;
            edits.Add(new(new(sheetId, position.Row, position.Column), content));
        }
        var copiedMerges = clipboard.MergedCells.Select(m => transpose
            ? new CellRectangle(output.Top + m.Left - clipboard.Range.Left, output.Left + m.Top - clipboard.Range.Top,
                output.Top + m.Right - clipboard.Range.Left, output.Left + m.Bottom - clipboard.Range.Top)
            : new CellRectangle(output.Top + m.Top - clipboard.Range.Top, output.Left + m.Left - clipboard.Range.Left,
                output.Top + m.Bottom - clipboard.Range.Top, output.Left + m.Right - clipboard.Range.Left)).ToArray();
        var targetMerges = sheet.MergedCells.Select(CellRectangle.Parse).Where(m => ValidationSafety.Overlaps(m, output)).ToHashSet();
        var copyMergeFormat = mode is CellPasteMode.All or CellPasteMode.Formats && !scalarMerge;
        var mergeChanged = copyMergeFormat && !targetMerges.SetEquals(copiedMerges);
        // Apply cells while unmerged so an Excel format-paste can retain covered values/formulas.
        var prepared = copyMergeFormat && targetMerges.Count > 0
            ? new MergeCellsCommand(sheetId, output, unmerge: true).Prepare(snapshot).Snapshot : snapshot;
        var draft = new SetCellsCommand(edits).Prepare(prepared);
        if (copyMergeFormat)
        {
            var merged = draft.Snapshot;
            foreach (var merge in copiedMerges)
                merged = new MergeCellsCommand(sheetId, merge) { PreserveCoveredContent = true }.Prepare(merged).Snapshot;
            if (!mergeChanged && draft.Changes.IsEmpty && !draft.StructureChanged)
                return new(snapshot, [], false, CalculationUnchanged: mode == CellPasteMode.Formats);
            draft = draft with { Snapshot = merged, StructureChanged = draft.StructureChanged || mergeChanged };
        }
        return mode == CellPasteMode.Formats ? draft with { CalculationUnchanged = !mergeChanged } : draft;
    }

    internal static string CopyFormula(WorkbookSnapshot snapshot, SheetId sourceSheet, SheetId targetSheet, string formula,
        CellPosition source, CellPosition target, bool transpose = false)
    {
        var expression = FormulaSyntax.Parse(formula);
        if (snapshot.FileState is not null && FormulaSyntax.Walk(FormulaBinding.Bind(expression, snapshot, sourceSheet, source))
            .OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name)))
            throw new NotSupportedException(MessageText.Get("导入公式含未支持函数，其复制引用语义尚未确认；请使用只贴值。"));
        if (sourceSheet != targetSheet && FormulaSyntax.Walk(expression).OfType<NameNode>().Any(n => n.Sheet is null &&
            snapshot.DefinedNames.Any(d => (d.Scope == sourceSheet || d.Scope == targetSheet) && d.Name.Equals(n.Name, StringComparison.OrdinalIgnoreCase))))
            throw new NotSupportedException(MessageText.Get("跨表复制涉及局部名称，尚不能安全确定目标作用域；整批操作未应用。"));
        var qualified = false;
        expression = FormulaSyntax.Map(expression, node =>
        {
            if (node is not StructuredReferenceNode { Table: null } r || StructuredReferences.Resolve(snapshot, sourceSheet, source, r) is not { } owner) return node;
            if (targetSheet == owner.Sheet.Id && owner.Table.Range.Contains(target.Row, target.Column)) return node;
            qualified = true; return r with { Table = owner.Table.Name };
        });
        return FormulaCopy.Translate(qualified ? "=" + FormulaSyntax.Print(expression) : formula, source, target, transpose);
    }
}
