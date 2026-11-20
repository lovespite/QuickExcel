using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public enum SheetNameConflictPolicy { Rename, Skip, Fail }

public sealed record DataTransferSheetSelection(SheetId SheetId, string? TargetName = null, CellRectangle? Range = null);
public sealed record DataTransferSheetPlan(SheetId SourceId, string SourceName, string? TargetName, bool Skipped);
public sealed record WorkbookDataTransferResult(WorkbookSnapshot Workbook, IReadOnlyList<string> Notices);

/// <summary>Explicit values-and-number-formats exchange; never carries executable source, references or attachments.</summary>
public static class WorkbookDataTransfer
{
    public const int MaximumTextCharacters = 32 * 1024 * 1024;
    public const string ValuesOnlyNotice = "仅传输单元格值和数字格式；公式转为结果，富文本转为纯文本，不包含批注、链接、图片、图表、表格定义、筛选、合并、布局、名称或脚本。源文件保持不变。";

    /// <summary>A selected rectangle is rebased to A1; a whole sheet retains its cell coordinates.</summary>
    public static WorkbookDataTransferResult ExtractValues(WorkbookSnapshot source,
        IReadOnlyList<DataTransferSheetSelection> selections, bool allowStaleCaches = false,
        CancellationToken cancellation = default)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(selections);
        cancellation.ThrowIfCancellationRequested();
        if (selections.Count is < 1 or > WorkbookLimits.MaxSheets || selections.Any(s => s is null) ||
            selections.Select(s => s.SheetId).Distinct().Count() != selections.Count)
            throw new ArgumentException(MessageText.Get("请选择至少一张工作表，每张工作表只能选择一次。"));
        var sheets = new List<WorksheetSnapshot>();
        var styles = new List<StyleDefinition> { new() };
        var styleIds = new Dictionary<StyleDefinition, int> { [styles[0]] = 0 };
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long characters = 0, count = 0;
        var stale = 0;
        foreach (var selected in selections)
        {
            cancellation.ThrowIfCancellationRequested();
            var sheet = source.GetSheet(selected.SheetId);
            var name = selected.TargetName ?? sheet.Name;
            WorkbookLimits.ValidateSheetName(name);
            if (!names.Add(name)) throw new ArgumentException(MessageText.Get("目标工作表名称重复，请修改名称。"));
            var range = selected.Range ?? new CellRectangle(0, 0, Math.Max(0, sheet.UsedRowCount - 1), Math.Max(0, sheet.UsedColumnCount - 1));
            ValidateRange(range);
            if ((count += range.Count) > WorkbookLimits.MaxCells)
                throw new InvalidOperationException(MessageText.Get("数据交换范围超过 1,000,000 格，请缩小范围。"));
            var cells = new Dictionary<CellPosition, CellSnapshot>();
            foreach (var (position, cell) in sheet.Cells)
            {
                cancellation.ThrowIfCancellationRequested();
                if (!range.Contains(position.Row, position.Column)) continue;
                var value = cell.Value;
                if (cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null)
                {
                    if (cell.Calculation?.Value is not { } result)
                        throw new InvalidOperationException(MessageText.Get("导入或导出区域中的公式缺少结果，请先重新计算；不能将缺失结果当作空白。"));
                    var isStale = source.FileState?.FormulaReadOnly == true || cell.Calculation.InputRevision != source.Revision ||
                        cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error);
                    if (isStale && !allowStaleCaches)
                        throw new InvalidOperationException(MessageText.Get("导入或导出区域存在过期或未验证的公式缓存，请先重新计算，或明确选择使用已有缓存。"));
                    if (isStale) stale++;
                    value = result;
                }
                if (value.Kind == CellValueKind.Text && (characters += value.TextValue.Length) > MaximumTextCharacters)
                    throw new InvalidOperationException(MessageText.Get("数据交换超过 32 MiB 字符预算，请缩小范围。"));
                var styleId = AddNumberFormat(styles, styleIds, source.Styles[sheet.GetStyleId(position.Row, position.Column)]);
                cells.Add(new(position.Row - range.Top, position.Column - range.Left), new(CellContent.FromLiteral(value, styleId)));
            }
            // A blank cell may inherit a date/number format from its row or column. Materialize
            // only those formats within the bounded exchange area, without copying layout metadata.
            if (sheet.Rows.Values.Any(r => r.ApplyStyle && HasNumberFormat(source.Styles[r.StyleId])) ||
                sheet.Columns.Values.Any(c => HasNumberFormat(source.Styles[c.StyleId])))
            {
                for (var row = range.Top; row <= range.Bottom; row++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    for (var column = range.Left; column <= range.Right; column++)
                    {
                        var position = new CellPosition(row - range.Top, column - range.Left);
                        if (cells.ContainsKey(position)) continue;
                        var style = source.Styles[sheet.GetInheritedStyleId(row, column)];
                        if (HasNumberFormat(style)) cells.Add(position,
                            new(CellContent.FromLiteral(CellValue.Blank, AddNumberFormat(styles, styleIds, style))));
                    }
                }
            }
            // Retain the selected rectangular extent, including trailing blank rows/columns.
            if (selected.Range.HasValue)
            {
                var corner = new CellPosition(range.Rows - 1, range.Columns - 1);
                if (!cells.ContainsKey(corner)) cells.Add(corner, new(CellContent.FromLiteral(CellValue.Blank,
                    AddNumberFormat(styles, styleIds, source.Styles[sheet.GetStyleId(range.Bottom, range.Right)]))));
            }
            sheets.Add(new(SheetId.New(), name, cells));
        }
        var notices = new List<string> { MessageText.Get(ValuesOnlyNotice) };
        if (stale > 0) notices.Add(MessageText.Format($"已明确使用 {stale} 个已有公式缓存；这些结果可能过期，未作为最新计算结果验证。"));
        if (source.FileState?.CompatibilityIssues.Length > 0)
            notices.Add(MessageText.Get("源工作簿包含兼容性诊断。仅提取可读取的单元格值，不保留未建模内容；请保留源文件。"));
        var workbook = new WorkbookSnapshot(Guid.NewGuid(), 0, source.DateSystem, sheets, styles: styles,
            scriptModules: ScriptModuleSet.Single("", ScriptModuleSet.CurrentApiVersion));
        return new(workbook, notices);
    }

    public static IReadOnlyList<DataTransferSheetPlan> PlanImport(WorkbookSnapshot target, WorkbookSnapshot sourceValues,
        SheetNameConflictPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(sourceValues);
        if (!Enum.IsDefined(policy)) throw new ArgumentOutOfRangeException(nameof(policy));
        var used = target.Sheets.Select(s => s.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var result = new List<DataTransferSheetPlan>();
        foreach (var sheet in sourceValues.Sheets)
        {
            var name = sheet.Name;
            if (used.Contains(name))
            {
                if (policy == SheetNameConflictPolicy.Skip)
                {
                    result.Add(new(sheet.Id, sheet.Name, null, true));
                    continue;
                }
                if (policy == SheetNameConflictPolicy.Fail)
                    throw new InvalidOperationException(MessageText.Format($"工作表名称“{name}”已存在，请修改名称或选择自动重命名。"));
                name = UniqueName(name, used);
            }
            used.Add(name);
            result.Add(new(sheet.Id, sheet.Name, name, false));
        }
        if (target.Sheets.Length + result.Count(p => !p.Skipped) > WorkbookLimits.MaxSheets)
            throw new InvalidOperationException(MessageText.Get("导入后工作表数量超过 64 张。"));
        return result;
    }

    internal static int AddNumberFormat(List<StyleDefinition> styles, Dictionary<StyleDefinition, int> ids, StyleDefinition source)
    {
        var format = new StyleDefinition { NumberFormatId = source.NumberFormatId, NumberFormatCode = source.NumberFormatCode };
        if (ids.TryGetValue(format, out var id)) return id;
        if (styles.Count >= WorkbookLimits.MaxStyles) throw new InvalidOperationException(MessageText.Get("目标样式数量超过上限。"));
        id = styles.Count;
        styles.Add(format);
        ids.Add(format, id);
        return id;
    }

    private static bool HasNumberFormat(StyleDefinition style) => style.NumberFormatId != 0 || style.NumberFormatCode is not null;

    private static void ValidateRange(CellRectangle range)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left);
        WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Bottom < range.Top || range.Right < range.Left) throw new ArgumentException(MessageText.Get("数据交换区域顺序无效。"));
    }

    private static string UniqueName(string name, HashSet<string> used)
    {
        for (var index = 2; ; index++)
        {
            var suffix = " (" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";
            var length = Math.Min(name.Length, WorkbookLimits.MaxSheetNameLength - suffix.Length);
            if (length > 0 && char.IsHighSurrogate(name[length - 1])) length--;
            var candidate = name[..length] + suffix;
            if (!used.Contains(candidate)) return candidate;
        }
    }
}

/// <summary>Adds a prepared values-only workbook as new sheets in one history transaction.</summary>
public sealed class ImportWorksheetDataCommand : DocumentCommand
{
    private readonly WorkbookSnapshot _source;
    private readonly SheetNameConflictPolicy _policy;

    public ImportWorksheetDataCommand(WorkbookSnapshot sourceValues, SheetNameConflictPolicy policy = SheetNameConflictPolicy.Rename)
    {
        ArgumentNullException.ThrowIfNull(sourceValues);
        if (!Enum.IsDefined(policy)) throw new ArgumentOutOfRangeException(nameof(policy));
        // Callers must first choose and acknowledge the explicit values-only projection.
        if (sourceValues.Sheets.Any(s => s.Cells.Values.Any(c => c.Content.Kind != CellContentKind.Literal || c.SpillOwner is not null || c.Content.RichText is not null) ||
            !s.MergedCells.IsEmpty || !s.Images.IsEmpty || s.Annotations.Links.Count > 0 || s.Annotations.Notes.Count > 0 ||
            !s.Annotations.NoteDrawing.IsEmpty || s.Rows.Count != 0 || s.Columns.Count != 0 || !s.Data.IsEmpty || !s.Printing.IsEmpty ||
            s.Visibility != SheetVisibility.Visible) || sourceValues.FileState is not null ||
            sourceValues.ScriptModules.Modules.Values.Any(m => m.Source.Length != 0) || !sourceValues.DefinedNames.IsEmpty ||
            sourceValues.PackageFiles.Entries.Count != 0 || sourceValues.LocalStorage.Entries.Count != 0)
            throw new ArgumentException(MessageText.Get("请先预览并确认仅提取单元格值，再导入工作表数据。"), nameof(sourceValues));
        if (sourceValues.Sheets.Sum(s => s.Cells.Values.Sum(c => c.Value.Kind == CellValueKind.Text ? (long)c.Value.TextValue.Length : 0)) > WorkbookDataTransfer.MaximumTextCharacters)
            throw new ArgumentException(MessageText.Get("数据交换超过 32 MiB 字符预算，请缩小范围。"), nameof(sourceValues));
        _source = sourceValues;
        _policy = policy;
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var plan = WorkbookDataTransfer.PlanImport(snapshot, _source, _policy);
        if (plan.All(p => p.Skipped)) return new(snapshot, [], false);
        if (snapshot.Sheets.Sum(s => (long)s.Cells.Count) + plan.Where(p => !p.Skipped).Sum(p => (long)_source.GetSheet(p.SourceId).Cells.Count) > WorkbookLimits.MaxCells)
            throw new InvalidOperationException(MessageText.Get("导入后工作簿超过 1,000,000 格。"));
        var styles = snapshot.Styles.ToList();
        var ids = new Dictionary<StyleDefinition, int>();
        for (var i = 0; i < styles.Count; i++) ids.TryAdd(styles[i], i);
        var sheets = snapshot.Sheets.ToList();
        var changes = ImmutableArray.CreateBuilder<CellChange>();
        foreach (var item in plan.Where(p => !p.Skipped))
        {
            var source = _source.GetSheet(item.SourceId);
            var sheetId = SheetId.New();
            var cells = new Dictionary<CellPosition, CellSnapshot>();
            foreach (var (position, cell) in source.Cells)
            {
                var style = _source.Styles[source.GetStyleId(position.Row, position.Column)];
                var value = cell.Value;
                if (_source.DateSystem != snapshot.DateSystem && value.Kind == CellValueKind.Number && CellDisplayFormatter.IsDateFormat(style, value.NumberValue) &&
                    !CellDisplayFormatter.IsTimeOnlyFormat(style, value.NumberValue))
                {
                    if (!new ExcelSerialDate(value.NumberValue, _source.DateSystem).TryConvertTo(snapshot.DateSystem, out var date))
                        throw new NotSupportedException(MessageText.Get("日期无法转换到目标日期系统（包括虚构的 1900-02-29）；整批导入未应用。"));
                    value = CellValue.FromNumber(date.Serial);
                }
                var copied = new CellSnapshot(CellContent.FromLiteral(value, WorkbookDataTransfer.AddNumberFormat(styles, ids, style)));
                cells.Add(position, copied);
                changes.Add(new(new(sheetId, position.Row, position.Column), null, copied));
            }
            sheets.Add(new(sheetId, item.TargetName!, cells));
        }
        var after = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, sheets,
            snapshot.ScriptSource, styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames,
            snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles, snapshot.LocalStorage);
        return new(after, changes.ToImmutable(), true);
    }
}
