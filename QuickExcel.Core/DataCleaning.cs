using System.Collections.Immutable;
using System.Text;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public enum CleaningColumnType { Text, Number, DateYmd, DateMdy, DateDmy, Skip }
public sealed record TextSplitOptions(string Delimiters = ",", bool CollapseDelimiters = false, char? Qualifier = '"',
    ImmutableArray<int> Breaks = default);
public sealed record NumericTextOptions(char DecimalSeparator = '.', char? GroupSeparator = ',', bool TrailingMinus = true);
public sealed record CleaningCellPreview(CellAddress Address, CellContent? Before, CellContent After);

/// <summary>An immutable, reviewable candidate. Commit cannot reinterpret newer inputs or expand its scope.</summary>
public sealed class CleaningPreview
{
    internal CleaningPreview(WorkbookSnapshot source, CommandDraft draft, IEnumerable<CellAddress> overwrites, IEnumerable<int> removed, IEnumerable<string> notices)
    {
        Source = source; Draft = draft; Overwrites = overwrites.ToImmutableArray(); RemovedRows = removed.ToImmutableArray(); Notices = notices.Distinct().ToImmutableArray();
        Changes = draft.Changes.Select(c => new CleaningCellPreview(c.Address, c.Before?.Content, c.After?.Content ?? CellContent.FromLiteral(CellValue.Blank))).ToImmutableArray();
    }
    internal WorkbookSnapshot Source { get; }
    internal CommandDraft Draft { get; }
    public ImmutableArray<CleaningCellPreview> Changes { get; }
    public ImmutableArray<CellAddress> Overwrites { get; }
    public ImmutableArray<int> RemovedRows { get; }
    public ImmutableArray<string> Notices { get; }
    public WorkbookSnapshot Result => Draft.Snapshot;
}

public sealed class ApplyCleaningCommand(CleaningPreview preview, bool allowOverwrite = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(preview);
        if (!SameInputs(snapshot, preview.Source)) throw new InvalidOperationException("清洗预览已过期，请重新预览；没有修改工作簿。");
        if (!allowOverwrite && !preview.Overwrites.IsEmpty) throw new InvalidOperationException("目标已有内容，请在预览中明确确认覆盖；没有修改工作簿。");
        // Navigation may change while a modal preview is open. Replay only the reviewed edits
        // against that current view; never restore the preview's old navigation or window.
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets, snapshot.ScriptSource,
            preview.Result.Styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion);
        return new SetCellsCommand(preview.Changes.Select(c => new CellEdit(c.Address, c.After))).Prepare(styled);
    }
    internal static bool SameInputs(WorkbookSnapshot a, WorkbookSnapshot b) => ReferenceEquals(a, b) ||
        a.DocumentId == b.DocumentId && a.Revision == b.Revision && a.DateSystem == b.DateSystem && a.TextCompatibilityVersion == b.TextCompatibilityVersion &&
        a.Styles == b.Styles && a.ScriptModules == b.ScriptModules && a.FileState == b.FileState && a.DefinedNames == b.DefinedNames && a.Sheets.Length == b.Sheets.Length &&
        a.Sheets.Zip(b.Sheets).All(p => p.First.Id == p.Second.Id && p.First.Name == p.Second.Name && ReferenceEquals(p.First.Cells, p.Second.Cells) &&
            ReferenceEquals(p.First.Rows, p.Second.Rows) && ReferenceEquals(p.First.Columns, p.Second.Columns) && p.First.MergedCells == p.Second.MergedCells &&
            p.First.Images == p.Second.Images && p.First.ShowGridLines == p.Second.ShowGridLines && p.First.DefaultRowHeight == p.Second.DefaultRowHeight &&
            p.First.TabColor == p.Second.TabColor && p.First.Visibility == p.Second.Visibility && p.First.Printing == p.Second.Printing &&
            p.First.Annotations == p.Second.Annotations && p.First.Data == p.Second.Data);
}

public static class DataCleaning
{
    public const int MaxCells = 20_000;
    private static WorksheetSnapshot RequireRange(WorkbookSnapshot book, SheetId sheetId, CellRectangle range)
    {
        if (book.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        WorkbookClipboard.ValidateRange(range); var sheet = book.GetSheet(sheetId);
        WorkbookClipboard.RequirePlainGeometry(sheet, range); WorkbookClipboard.RequireVisibleDimensions(sheet, range);
        if (sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => range.Contains(p.Row, p.Column)) ||
            sheet.Cells.Any(p => range.Contains(p.Key.Row, p.Key.Column) && p.Value.Content.RichText is not null))
            throw new NotSupportedException("清洗区域含富文本、链接或批注，请先明确处理这些内容；整批操作未应用。");
        var hidden = WorksheetFilterEngine.Evaluate(book, sheetId);
        if (hidden.Any(r => r >= range.Top && r <= range.Bottom)) throw new InvalidOperationException("清洗区域含筛选隐藏行，请先清除筛选；不跨隐藏行映射记录。");
        return sheet;
    }
    public static CleaningPreview RemoveDuplicates(WorkbookSnapshot book, SheetId sheetId, CellRectangle range, IEnumerable<int> keyColumns, bool hasHeader = true)
    {
        var sheet = RequireRange(book, sheetId, range); var keys = keyColumns.Take(1025).ToArray();
        if (WorksheetOutline.HasGroups(sheet)) throw new NotSupportedException("删除重复项暂不压缩分组，请先取消分组。");
        if (keys.Length is 0 or > 1024 || keys.Any(c => c < range.Left || c > range.Right) || keys.Distinct().Count() != keys.Length)
            throw new ArgumentException("请选择不重复的区域内判重列。");
        if (sheet.Data.Filter is { } filter && ValidationSafety.Overlaps(filter.Range, range) || sheet.Data.Tables.Any(t => ValidationSafety.Overlaps(t.Range, range)) ||
            sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, range))) ||
            sheet.Data.Validations.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, range))))
            throw new NotSupportedException("删除重复项暂不压缩筛选、Table、条件格式或验证区域；请先复制值到普通区域再清洗。");
        StructureEdits.RequireTransformableFunctions(book);
        var first = range.Top + (hasHeader ? 1 : 0); var keep = new List<int>(); var removed = new List<int>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); long characters = 0;
        for (var row = first; row <= range.Bottom; row++)
        {
            var key = new StringBuilder();
            foreach (var column in keys)
            {
                var cell = sheet.GetCell(row, column);
                if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation is not { State: CalculationState.Fresh or CalculationState.Error, Value: not null } || cell.Calculation.InputRevision != book.Revision))
                    throw new InvalidOperationException("判重列含待重算公式，请先完成重算并重新预览。");
                var value = cell?.Value ?? CellValue.Blank;
                var text = WorksheetFilterEngine.ValueLabel(value, TablePresentation.NumberStyle(sheet, book.Styles, row, column), book.DateSystem);
                var kind = value.Kind switch { CellValueKind.Number => "N", CellValueKind.Blank => "B", CellValueKind.Error => "E", _ => "T" };
                key.Append(kind).Append(text.Length).Append(':').Append(text);
            }
            characters += key.Length; if (characters > 4_000_000) throw new InvalidOperationException("判重文本超过 4,000,000 字符，请缩小区域。");
            if (seen.Add(key.ToString())) keep.Add(row); else removed.Add(row);
        }
        var edits = new List<CellEdit>();
        if (removed.Count == 0) return new(book, new(book, [], false), [], [], ["没有找到重复项，数据保持不变。"]);
        for (var row = first; row <= range.Bottom; row++)
        for (var col = range.Left; col <= range.Right; col++)
        {
            var index = row - first; CellContent content;
            if (index >= keep.Count) content = CellContent.FromLiteral(CellValue.Blank);
            else
            {
                var oldRow = keep[index]; content = sheet.GetCell(oldRow, col)?.Content ?? CellContent.FromLiteral(CellValue.Blank, sheet.GetStyleId(oldRow, col));
                if (oldRow != row) content = content.WithStyle(sheet.GetStyleId(oldRow, col));
                if (oldRow != row && content.Kind == CellContentKind.Formula)
                    content = CellContent.FromFormula(PasteCellsCommand.CopyFormula(book, sheetId, sheetId, content.FormulaText, new(oldRow, col), new(row, col)), content.StyleId);
            }
            edits.Add(new(new(sheetId, row, col), content));
        }
        return new(book, new SetCellsCommand(edits).Prepare(book), [], removed,
            ["保留首次出现的记录，仅在选区内向上压缩并清空尾部；整张工作表行数及区外单元格位置不变。", "按所选列的值类型和基础显示格式判重；不区分文本大小写，不自动合并文本 ID 与数值。"]);
    }
    public static CleaningPreview SplitText(WorkbookSnapshot book, SheetId sheetId, CellRectangle source, CellPosition destination,
        TextSplitOptions options, IEnumerable<CleaningColumnType>? columnTypes = null, NumericTextOptions? numbers = null)
    {
        var sheet = RequireRange(book, sheetId, source);
        if (source.Left != source.Right) throw new ArgumentException("分列来源须为单列矩形。");
        var types = (columnTypes ?? []).Take(257).ToArray();
        if (types.Length > 256 || types.Any(t => !Enum.IsDefined(t))) throw new ArgumentException("分列类型列表最多 256 列。");
        var parsed = new List<string[]>(); var maximum = 0; long characters = 0;
        for (var row = source.Top; row <= source.Bottom; row++)
        {
            var cell = sheet.GetCell(row, source.Left);
            if (cell is not null && (cell.Content.Kind != CellContentKind.Literal || cell.Value.Kind is not (CellValueKind.Text or CellValueKind.Blank)))
                throw new InvalidOperationException($"第 {row + 1} 行不是普通文本；分列不隐式转换数值、日期或公式。");
            var text = cell?.Value.Kind == CellValueKind.Text ? cell.Value.TextValue : ""; characters += text.Length;
            if (characters > 4_000_000) throw new InvalidOperationException("分列来源超过 4,000,000 字符。");
            var fields = CleaningText.Parse(text, options); parsed.Add(fields); maximum = Math.Max(maximum, fields.Length);
        }
        if (maximum > 256) throw new InvalidOperationException("分列结果超过 256 列。");
        CleaningColumnType Type(int i) => i < types.Length ? types[i] : CleaningColumnType.Text;
        var count = Enumerable.Range(0, maximum).Count(i => Type(i) != CleaningColumnType.Skip);
        if (count == 0) throw new ArgumentException("至少保留一个输出列。");
        var output = new CellRectangle(destination.Row, destination.Column, checked(destination.Row + parsed.Count - 1), checked(destination.Column + count - 1));
        RequireRange(book, sheetId, output);
        var builder = new Builder(book, sheet); var overwrites = new List<CellAddress>();
        for (var r = 0; r < parsed.Count; r++)
        {
            var targetColumn = destination.Column;
            for (var c = 0; c < maximum; c++)
            {
                var type = Type(c); if (type == CleaningColumnType.Skip) continue;
                var position = new CellPosition(destination.Row + r, targetColumn++);
                var text = c < parsed[r].Length ? parsed[r][c] : "";
                builder.Convert(position, text, type, numbers ?? new());
                var old = sheet.GetCell(position.Row, position.Column);
                if (!source.Contains(position.Row, position.Column) && old is not null && (old.Content.Kind == CellContentKind.Formula || old.Value.Kind != CellValueKind.Blank))
                    overwrites.Add(new(sheetId, position.Row, position.Column));
            }
        }
        return builder.Finish(overwrites);
    }
    public static CleaningPreview ConvertTextNumbers(WorkbookSnapshot book, SheetId sheetId, CellRectangle range, NumericTextOptions? numbers = null)
    {
        var sheet = RequireRange(book, sheetId, range); var builder = new Builder(book, sheet); var skipped = 0;
        for (var row = range.Top; row <= range.Bottom; row++)
        for (var col = range.Left; col <= range.Right; col++)
        {
            var cell = sheet.GetCell(row, col);
            if (cell?.Content.Kind != CellContentKind.Literal || cell.Value.Kind != CellValueKind.Text) { skipped++; continue; }
            builder.Convert(new(row, col), cell.Value.TextValue, CleaningColumnType.Number, numbers ?? new());
        }
        if (skipped > 0) builder.Notices.Add($"保留 {skipped} 个非文本单元格（包括公式），不转换其值。");
        return builder.Finish([]);
    }
    private sealed class Builder(WorkbookSnapshot book, WorksheetSnapshot sheet)
    {
        private readonly List<StyleDefinition> _styles = book.Styles.ToList();
        private readonly List<CellEdit> _edits = [];
        public List<string> Notices { get; } = ["未明确指定类型的列默认保留文本；不猜测日期或把等号开头的文本执行为公式。"];
        public void Convert(CellPosition position, string text, CleaningColumnType type, NumericTextOptions numbers)
        {
            (CellValue Value, int FormatId, string? FormatCode) converted;
            try { converted = CleaningText.Convert(text, type, numbers, book.DateSystem, Notices); }
            catch (Exception ex) when (ex is ArgumentException or OverflowException)
            { throw new ArgumentException($"目标第 {position.Row + 1} 行、第 {position.Column + 1} 列：{ex.Message}", ex); }
            var baseStyle = book.Styles[sheet.GetStyleId(position.Row, position.Column)];
            var style = baseStyle with { NumberFormatId = converted.FormatId, NumberFormatCode = converted.FormatCode };
            var id = _styles.IndexOf(style);
            if (id < 0) { if (_styles.Count >= 4096) throw new InvalidOperationException("样式超过 4096，清洗未应用。"); id = _styles.Count; _styles.Add(style); }
            _edits.Add(new(new(sheet.Id, position.Row, position.Column), CellContent.FromLiteral(converted.Value, id)));
        }
        public CleaningPreview Finish(IEnumerable<CellAddress> overwrites)
        {
            var styled = new WorkbookSnapshot(book.DocumentId, book.Revision, book.DateSystem, book.Sheets, book.ScriptSource, _styles, book.ScriptModules,
                book.FileState, book.DefinedNames, book.ActiveSheetId, book.Window, book.TextCompatibilityVersion);
            var draft = new SetCellsCommand(_edits).Prepare(styled);
            if (draft.Changes.IsEmpty) draft = new(book, [], false);
            return new(book, draft, overwrites, [], Notices);
        }
    }
}
