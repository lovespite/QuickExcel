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
            return new CellSnapshot(cell.Content.WithStyle(sheet.GetStyleId(row, column)), cell.Calculation);
        })).ToImmutableArray();
        HasAnnotations = sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => range.Contains(p.Row, p.Column));
        HasConditionalFormats = sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, range))) ||
            snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase)) == true;
        HasOldResults = Cells.Any(c => c.Content.Kind == CellContentKind.Formula && (c.Calculation?.Value is null ||
            c.Calculation.InputRevision != snapshot.Revision || c.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)));
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

    public static WorkbookClipboard Capture(WorkbookSnapshot snapshot, SheetId sheetId, CellRectangle range, bool cut = false)
    {
        ValidateRange(range); var sheet = snapshot.GetSheet(sheetId);
        RequirePlainGeometry(sheet, range);
        RequireVisibleDimensions(sheet, range);
        return new(snapshot, sheet, range, cut);
    }
    internal static void ValidateRange(CellRectangle range)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Top > range.Bottom || range.Left > range.Right || (long)(range.Bottom - range.Top + 1) * (range.Right - range.Left + 1) > MaxCells)
            throw new ArgumentException("内部复制/剪切需要不超过 20,000 格的连续矩形。");
    }
    internal static void RequirePlainGeometry(WorksheetSnapshot sheet, CellRectangle range)
    {
        bool Intersects(int top, int left, int bottom, int right) => top <= range.Bottom && bottom >= range.Top && left <= range.Right && right >= range.Left;
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(m => Intersects(m.Top, m.Left, m.Bottom, m.Right)) ||
            sheet.Images.Any(i => Intersects(i.RowIndex, i.ColumnIndex, i.RowIndex + i.RowSpan - 1, i.ColumnIndex + i.ColumnSpan - 1)))
            throw new NotSupportedException("内部复制/粘贴区域含合并单元格或图片，尚不能完整复制其结构；整批操作未应用。");
    }
    internal static void RequireVisibleDimensions(WorksheetSnapshot sheet, CellRectangle range)
    {
        if (sheet.Rows.Any(r => r.Key >= range.Top && r.Key <= range.Bottom && r.Value.Hidden) ||
            sheet.Columns.Any(c => c.Key >= range.Left && c.Key <= range.Right && c.Value.Hidden))
            throw new InvalidOperationException("内部复制/剪切/粘贴区域含隐藏行列。请先取消隐藏；整批操作未应用。");
    }
}

/// <summary>One atomic paste. The caller advances its lease only after its own successful paste; any other edit expires it.</summary>
public sealed class PasteCellsCommand(WorkbookClipboard clipboard, SheetId sheetId, CellRectangle target,
    long expectedRevision, CellPasteMode mode = CellPasteMode.All, bool transpose = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(clipboard);
        if (!Enum.IsDefined(mode)) throw new ArgumentException("未知粘贴模式。");
        if (snapshot.DocumentId != clipboard.DocumentId || snapshot.Revision != expectedRevision || expectedRevision < clipboard.Revision)
            throw new InvalidOperationException("复制后文档已改变，请重新复制或剪切；整批操作未应用。");
        WorkbookClipboard.ValidateRange(target);
        var rows = transpose ? clipboard.Columns : clipboard.Rows; var columns = transpose ? clipboard.Rows : clipboard.Columns;
        var single = target.Top == target.Bottom && target.Left == target.Right;
        var fill = clipboard.Rows == 1 && clipboard.Columns == 1 && !clipboard.IsCut;
        if (fill) { rows = target.Bottom - target.Top + 1; columns = target.Right - target.Left + 1; }
        if (!single && !fill && (target.Bottom - target.Top + 1 != rows || target.Right - target.Left + 1 != columns))
            throw new ArgumentException("选区尺寸须与粘贴数据相同，或只选择左上角单元格。");
        var output = new CellRectangle(target.Top, target.Left, checked(target.Top + rows - 1), checked(target.Left + columns - 1));
        WorkbookClipboard.ValidateRange(output);
        var sheet = snapshot.GetSheet(sheetId);
        WorkbookClipboard.RequirePlainGeometry(sheet, output); WorkbookClipboard.RequireVisibleDimensions(sheet, output);
        if (mode is CellPasteMode.All or CellPasteMode.Formats && (clipboard.HasConditionalFormats ||
            sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, output)))))
            throw new NotSupportedException("完整/格式粘贴涉及条件格式，尚不能复制或拆分其作用范围；可明确选择只贴值或公式，规则可在条件格式管理器中编辑。");
        if (clipboard.IsCut)
        {
            if (mode != CellPasteMode.All || transpose) throw new NotSupportedException("剪切仅支持完整粘贴，不支持选择性粘贴或转置；源区域尚未修改。");
            if (expectedRevision != clipboard.Revision) throw new InvalidOperationException("剪切后文档已改变，请重新剪切。");
            var source = snapshot.GetSheet(clipboard.SheetId);
            for (var r = 0; r < rows; r++)
            for (var c = 0; c < columns; c++)
                if (!(source.Id == sheet.Id && clipboard.Range.Contains(output.Top + r, output.Left + c)) &&
                    sheet.GetCell(output.Top + r, output.Left + c)?.Content.RichText is not null &&
                    clipboard.Cells[r * clipboard.Columns + c].Content.RichText is null)
                    throw new InvalidOperationException("剪切目标包含富文本，请先明确转换为普通文本；源区域尚未修改。");
            // MoveCells preserves explicit styles. Inherited axis styles need separate move semantics.
            if (Enumerable.Range(clipboard.Range.Top, clipboard.Rows).Any(r => source.Rows.GetValueOrDefault(r)?.ApplyStyle == true) ||
                Enumerable.Range(clipboard.Range.Left, clipboard.Columns).Any(c => source.Columns.GetValueOrDefault(c)?.StyleId > 0) ||
                Enumerable.Range(output.Top, rows).Any(r => sheet.Rows.GetValueOrDefault(r)?.ApplyStyle == true) ||
                Enumerable.Range(output.Left, columns).Any(c => sheet.Columns.GetValueOrDefault(c)?.StyleId > 0))
                throw new NotSupportedException("含整行/整列继承样式的区域暂不支持剪切；请使用复制。");
            return new MoveCellsCommand(clipboard.SheetId, new(clipboard.Range.Top, clipboard.Range.Left),
                new(clipboard.Range.Bottom, clipboard.Range.Right), sheetId, new(output.Top, output.Left)).Prepare(snapshot);
        }
        if (mode == CellPasteMode.All && (clipboard.HasAnnotations ||
            sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => output.Contains(p.Row, p.Column))))
            throw new NotSupportedException("完整粘贴涉及超链接/批注，尚不能安全变换附件；可明确选择只贴值、公式或格式。");
        if (mode == CellPasteMode.Values && clipboard.Cells.Any(c => c.Content.Kind == CellContentKind.Formula && c.Calculation?.Value is null))
            throw new InvalidOperationException("源公式缺少结果，不能把它粘贴为空白；请先完成重算。");
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
                    throw new NotSupportedException("富文本分段格式不能通过只贴格式准确映射；请先明确转换为普通文本，整批操作未应用。");
                content = current.WithStyle(style);
            }
            else if (mode == CellPasteMode.Values) content = CellContent.FromLiteral(original.Value, style);
            else if (original.Content.Kind == CellContentKind.Formula)
            {
                content = CellContent.FromFormula(CopyFormula(snapshot, clipboard.SheetId, sheetId, original.Content.FormulaText,
                    new(clipboard.Range.Top + sr, clipboard.Range.Left + sc), position, transpose), style);
            }
            else content = mode == CellPasteMode.All ? original.Content.WithStyle(style) : CellContent.FromLiteral(original.Value, style);
            edits.Add(new(new(sheetId, position.Row, position.Column), content));
        }
        var draft = new SetCellsCommand(edits).Prepare(snapshot);
        return mode == CellPasteMode.Formats ? draft with { CalculationUnchanged = true } : draft;
    }

    internal static string CopyFormula(WorkbookSnapshot snapshot, SheetId sourceSheet, SheetId targetSheet, string formula,
        CellPosition source, CellPosition target, bool transpose = false)
    {
        var expression = FormulaSyntax.Parse(formula);
        if (snapshot.FileState is not null && FormulaSyntax.Walk(FormulaBinding.Bind(expression, snapshot, sourceSheet, source))
            .OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name)))
            throw new NotSupportedException("导入公式含未支持函数，其复制引用语义尚未确认；请使用只贴值。");
        if (sourceSheet != targetSheet && FormulaSyntax.Walk(expression).OfType<NameNode>().Any(n => n.Sheet is null &&
            snapshot.DefinedNames.Any(d => (d.Scope == sourceSheet || d.Scope == targetSheet) && d.Name.Equals(n.Name, StringComparison.OrdinalIgnoreCase))))
            throw new NotSupportedException("跨表复制涉及局部名称，尚不能安全确定目标作用域；整批操作未应用。");
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
