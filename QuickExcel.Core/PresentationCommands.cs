using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Stable rectangular row sort. Moved formulas use copy semantics; outside references keep their addresses.</summary>
public sealed class SortRangeCommand : DocumentCommand
{
    public const int MaxCells = 20_000;
    private readonly SortDescriptor _descriptor;
    internal SortDescriptor Descriptor => _descriptor;
    public SortRangeCommand(SheetId sheetId, CellPosition first, CellPosition last, int keyColumn, bool descending = false)
        : this(new SortDescriptor(sheetId, first, last, [new SortKey(keyColumn, descending)])) { }

    public SortRangeCommand(SortDescriptor descriptor) => _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));

    /// <summary>Excel-style “扩展列”: keep extending across adjacent columns that hold record data, excluding the header row.</summary>
    public static CellRectangle ExpandColumns(WorksheetSnapshot sheet, CellRectangle range, bool hasHeader)
    {
        var top = range.Top + (hasHeader ? 1 : 0);
        if (top > range.Bottom) return range;
        var rows = range.Bottom - top + 1;
        static bool HasRecordCell(WorksheetSnapshot sheet, int row, int column)
        {
            var cell = sheet.GetCell(row, column);
            return cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.Value.Kind != CellValueKind.Blank);
        }
        bool ColumnHasRecordData(int column) => column >= 0 && column < WorkbookLimits.MaxColumns &&
            Enumerable.Range(top, rows).Any(row => HasRecordCell(sheet, row, column));
        void CheckBudget(int left, int right)
        {
            if ((long)rows * (right - left + 1) > MaxCells)
                throw new InvalidOperationException($"扩展列后排序范围超过 {MaxCells:N0} 格，请先缩小选区或关闭“扩展列”。");
        }
        var first = range.Left; var last = range.Right;
        while (first > 0 && ColumnHasRecordData(first - 1)) { first--; CheckBudget(first, last); }
        while (last < WorkbookLimits.MaxColumns - 1 && ColumnHasRecordData(last + 1)) { last++; CheckBudget(first, last); }
        return range with { Left = first, Right = last };
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        StructureEdits.RequireTransformableFunctions(snapshot);
        var (sheetId, first, last) = (_descriptor.SheetId, _descriptor.First, _descriptor.Last);
        WorkbookLimits.ValidatePosition(first.Row, first.Column);
        WorkbookLimits.ValidatePosition(last.Row, last.Column);
        if (first.Row > last.Row || first.Column > last.Column || _descriptor.Keys.IsEmpty || _descriptor.Keys.Length > 32 ||
            _descriptor.Keys.Any(key => key.Column < first.Column || key.Column > last.Column) ||
            _descriptor.Keys.Select(key => (key.Column, key.Basis)).Distinct().Count() != _descriptor.Keys.Length ||
            (long)(last.Row - first.Row + 1) * (last.Column - first.Column + 1) > MaxCells)
            throw new ArgumentException($"排序需要不超过 {MaxCells:N0} 格的矩形，以及 1–32 个不重复的区域内排序键。");
        foreach (var key in _descriptor.Keys) key.Validate();
        var sheet = snapshot.GetSheet(sheetId);
        StructureEdits.RequireSimpleStructure(sheet);
        if (WorksheetOutline.HasGroups(sheet)) throw new NotSupportedException("含分组的工作表暂不排序，请先取消分组。");
        if (_descriptor.ExpandColumns)
        {
            // Keys stay in the explicitly selected columns; only the moved record becomes wider.
            var expanded = ExpandColumns(sheet, new CellRectangle(first.Row, first.Column, last.Row, last.Column), _descriptor.HasHeader);
            first = first with { Column = expanded.Left }; last = last with { Column = expanded.Right };
        }
        var dataFirst = first.Row + (_descriptor.HasHeader ? 1 : 0);
        if (sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a =>
                ValidationSafety.Overlaps(a, new(dataFirst, first.Column, last.Row, last.Column)) &&
                (a.Top > dataFirst || a.Bottom < last.Row))) ||
            !sheet.Data.ConditionalFormats.IsEmpty && _descriptor.Keys.Any(k => k.Basis != SortBasis.Value))
            throw new NotSupportedException("条件格式仅支持规则覆盖全部待排序行的按值排序；部分行区域或条件颜色排序尚未支持。");
        var affectedTables = sheet.Data.Tables.Where(t => first.Row <= t.Range.Bottom && last.Row >= t.Range.Top && first.Column <= t.Range.Right && last.Column >= t.Range.Left).ToArray();
        var table = affectedTables.SingleOrDefault();
        if (table is not null && _descriptor.Keys.Any(k => k.Basis != SortBasis.Value)) throw new NotSupportedException("Table 的条纹/差异颜色排序暂不支持重应用；请选择按值排序。");
        if (table is not null && (first.Column != table.Range.Left || last.Column != table.Range.Right || dataFirst != table.Range.Top + (table.HeaderRow ? 1 : 0) || last.Row != table.Range.Bottom - (table.TotalsRow ? 1 : 0)))
            throw new InvalidOperationException("表排序必须覆盖完整数据区，保留表头和汇总行；本次没有修改。");
        (CellValue Value, StyleDefinition Style) Key(int row, SortKey key)
        {
            var cell = sheet.GetCell(row, key.Column);
            if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation?.Value is null ||
                cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.InputRevision != snapshot.Revision))
                throw new InvalidOperationException("排序关键列包含旧缓存或缺少结果，请先完成重算。");
            return (cell?.Value ?? CellValue.Blank, snapshot.Styles[sheet.GetStyleId(row, key.Column)]);
        }
        // Materialize every key before sorting: even a later key that never breaks a tie must be fresh.
        var records = Enumerable.Range(dataFirst, Math.Max(0, last.Row - dataFirst + 1))
            .Select(row => (Row: row, Keys: _descriptor.Keys.Select(key => Key(row, key)).ToArray())).ToArray();
        var order = records.OrderBy(record => record, Comparer<(int Row, (CellValue Value, StyleDefinition Style)[] Keys)>.Create((a, b) =>
        {
            for (var i = 0; i < _descriptor.Keys.Length; i++)
            {
                var comparison = _descriptor.Keys[i].Compare(a.Keys[i], b.Keys[i]);
                if (comparison != 0) return comparison;
            }
            return a.Row.CompareTo(b.Row);
        })).Select(record => record.Row).ToArray();
        var sortState = dataFirst <= last.Row ? new WorksheetSortState(new(dataFirst, first.Column, last.Row, last.Column), _descriptor.Keys) : null;
        var dataState = table is null ? sheet.Data with { Sort = sortState } : sheet.Data with { Tables = sheet.Data.Tables.Replace(table, table.WithSort(sortState)) };
        if (order.Select((row, index) => row == dataFirst + index).All(v => v))
            return sheet.Data.ContentEquals(dataState) ? new(snapshot, [], false) :
                new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(dataState) : s)), [], false) { CalculationUnchanged = true };
        var cells = sheet.Cells.ToBuilder();
        for (var index = 0; index < order.Length; index++)
        for (var column = first.Column; column <= last.Column; column++)
        {
            var row = dataFirst + index;
            var source = sheet.GetCell(order[index], column);
            var position = new CellPosition(row, column);
            if (source is null) { cells.Remove(position); continue; }
            if (source.Content.Kind == CellContentKind.Formula && order[index] != row)
            {
                var delta = row - order[index];
                FormulaNode Shift(FormulaNode node)
                {
                    if (node is not ReferenceNode r || r.WholeColumns) return node;
                    var start = r.Row + (r.RowAbsolute ? 0 : delta);
                    var end = r.IsRange ? r.EndRow + (r.EndRowAbsolute ? 0 : delta) : start;
                    if (start < 0 || end < 0 || start >= WorkbookLimits.MaxRows || end >= WorkbookLimits.MaxRows)
                        return new ValueNode(CellValue.FromError(CellErrorCode.Reference));
                    return r with { Row = start, EndRow = end };
                }
                source = new(CellContent.FromFormula("=" + FormulaSyntax.Print(FormulaSyntax.Map(
                    FormulaSyntax.Parse(source.Content.FormulaText), Shift)), source.Content.StyleId));
            }
            cells[position] = source;
        }
        return StructureEdits.Draft(snapshot, snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithCells(cells.ToImmutable()).WithData(dataState) : s)));
    }

    // Numbers, ordinal-ignore-case text, booleans, errors, blanks. Blanks stay last in both directions.
    internal static int Compare(CellValue a, CellValue b, bool descending, bool caseSensitive = false)
    {
        static int Rank(CellValue v) => v.Kind switch { CellValueKind.Number => 0, CellValueKind.Text => 1, CellValueKind.Boolean => 2, CellValueKind.Error => 3, _ => 4 };
        if (a.Kind == CellValueKind.Blank || b.Kind == CellValueKind.Blank) return Rank(a).CompareTo(Rank(b));
        var comparison = Rank(a).CompareTo(Rank(b));
        if (comparison == 0) comparison = a.Kind switch
        {
            CellValueKind.Number => a.NumberValue.CompareTo(b.NumberValue),
            CellValueKind.Text => (caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase).Compare(a.TextValue, b.TextValue),
            CellValueKind.Boolean => a.BooleanValue.CompareTo(b.BooleanValue),
            CellValueKind.Error => a.ErrorCode.CompareTo(b.ErrorCode), _ => 0
        };
        return descending ? -comparison : comparison;
    }
}

public sealed record CellStylePatch(bool? Bold = null, bool? Italic = null, string? FontColor = null,
    string? FillColor = null, int? Horizontal = null, int? NumberFormatId = null, string? NumberFormatCode = null,
    bool? Border = null, double? FontSize = null, bool Reset = false,
    string? FontName = null, int? Vertical = null, bool? WrapText = null,
    bool ClearFontColor = false, bool ClearFillColor = false)
{
    internal StyleDefinition Apply(StyleDefinition style)
    {
        if (Reset) return new();
        if (FontName is not null && (string.IsNullOrWhiteSpace(FontName) || FontName.Length > 256))
            throw new ArgumentException("字体名称不能为空且不能超过 256 个字符。");
        if (FontSize is <= 0 or > 409 || FontSize.HasValue && !double.IsFinite(FontSize.Value))
            throw new ArgumentException("字号必须大于 0 且不超过 409。");
        if (Vertical is < 0 or > 3) throw new ArgumentException("不支持的垂直对齐方式。");
        if (NumberFormatId is < 0 or >= 164) throw new ArgumentException("自定义数字格式使用代码和规范化 ID 0，内置格式 ID 必须小于 164。");
        static string? Color(string? color)
        {
            if (color is not null && (color.Length != 8 || color.Any(c => !Uri.IsHexDigit(c))))
                throw new ArgumentException("颜色必须是 8 位 ARGB 十六进制。");
            return color;
        }
        return style with
        {
            Bold = Bold ?? style.Bold, Italic = Italic ?? style.Italic,
            FontColor = ClearFontColor ? null : Color(FontColor) ?? style.FontColor,
            FillColor = ClearFillColor ? null : Color(FillColor) ?? style.FillColor,
            FontSize = FontSize ?? style.FontSize, Horizontal = Horizontal ?? style.Horizontal,
            FontName = FontName ?? style.FontName, Vertical = Vertical ?? style.Vertical,
            FontScheme = FontName is null ? style.FontScheme : null,
            EastAsianFontName = FontName is null ? style.EastAsianFontName : null,
            WrapText = WrapText ?? style.WrapText,
            NumberFormatId = NumberFormatId ?? style.NumberFormatId,
            NumberFormatCode = NumberFormatId.HasValue ? NumberFormatCode : style.NumberFormatCode,
            LeftBorder = Border.HasValue ? (Border.Value ? 1 : 0) : style.LeftBorder,
            RightBorder = Border.HasValue ? (Border.Value ? 1 : 0) : style.RightBorder,
            TopBorder = Border.HasValue ? (Border.Value ? 1 : 0) : style.TopBorder,
            BottomBorder = Border.HasValue ? (Border.Value ? 1 : 0) : style.BottomBorder
        };
    }
}

public sealed class SetCellStyleCommand : DocumentCommand
{
    private readonly ImmutableArray<CellAddress> _addresses;
    private readonly CellStylePatch _patch;
    public SetCellStyleCommand(IEnumerable<CellAddress> addresses, CellStylePatch patch)
    {
        _addresses = addresses.Take(20_001).ToImmutableArray();
        if (_addresses.Length > 20_000) throw new ArgumentException("样式编辑上限为 20,000 格。");
        _patch = patch;
    }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var styles = snapshot.Styles.ToList();
        var styleIds = new Dictionary<StyleDefinition, int>();
        for (var i = 0; i < styles.Count; i++) styleIds.TryAdd(styles[i], i);
        var edits = new List<CellEdit>();
        foreach (var address in _addresses.Distinct())
        {
            var sheet = snapshot.GetSheet(address.SheetId);
            var style = _patch.Apply(snapshot.Styles[sheet.GetStyleId(address.Row, address.Column)]);
            if (!styleIds.TryGetValue(style, out var id))
            {
                if (styles.Count >= WorkbookLimits.MaxStyles) throw new InvalidOperationException($"样式数量超过 {WorkbookLimits.MaxStyles}，操作未应用。");
                id = styles.Count; styles.Add(style); styleIds.Add(style, id);
            }
            var content = sheet.GetCell(address.Row, address.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank);
            if (content.RichText is { } rich && (_patch.Reset || _patch.Bold.HasValue || _patch.Italic.HasValue || _patch.FontName is not null || _patch.FontSize.HasValue || _patch.FontColor is not null || _patch.ClearFontColor))
            {
                // Whole-cell font commands explicitly override every run, like Excel; other patches leave runs intact.
                var runs = rich.Runs.Select(run =>
                {
                    if (_patch.Reset) return run with { Font = null };
                    var inherited = snapshot.Styles[content.StyleId];
                    var font = run.Font ?? new TextRunFont(inherited.FontName, inherited.FontSize, inherited.Bold, inherited.Italic, inherited.FontColor,
                        Scheme: inherited.FontScheme, EastAsianName: inherited.EastAsianFontName, ResolvedName: inherited.FontName,
                        Underline: inherited.Underline, Strike: inherited.Strike);
                    var clearColor = _patch.ClearFontColor;
                    return run with { Font = font with
                    {
                        Bold = _patch.Bold ?? font.Bold, Italic = _patch.Italic ?? font.Italic, Size = _patch.FontSize ?? font.Size,
                        Name = _patch.FontName ?? font.Name, ResolvedName = _patch.FontName ?? font.ResolvedName,
                        EastAsianName = _patch.FontName is null ? font.EastAsianName : null, Scheme = _patch.FontName is null ? font.Scheme : null,
                        Color = clearColor ? null : _patch.FontColor ?? font.Color, RawRgb = clearColor ? null : _patch.FontColor ?? font.RawRgb,
                        ThemeColor = clearColor || _patch.FontColor is not null ? null : font.ThemeColor,
                        IndexedColor = clearColor || _patch.FontColor is not null ? null : font.IndexedColor,
                        AutoColor = clearColor || _patch.FontColor is not null ? null : font.AutoColor,
                        Tint = clearColor || _patch.FontColor is not null ? null : font.Tint
                    } };
                });
                content = CellContent.FromRichText(new(runs), id);
            }
            edits.Add(new(address, content.WithStyle(id)));
        }
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets,
            snapshot.ScriptSource, styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles);
        var draft = new SetCellsCommand(edits).Prepare(styled);
        return draft.Changes.IsEmpty ? new(snapshot, [], false) : draft with { CalculationUnchanged = true };
    }
}

public sealed class SetAxisMetadataCommand(SheetId sheetId, SheetAxis axis, int first, int count, double? size = null, bool? hidden = null) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (!Enum.IsDefined(axis) || first < 0 || count < 1 || first > limit - count) throw new ArgumentOutOfRangeException(nameof(first));
        WorkbookLimits.ValidateDimension(size, nameof(size));
        var sheet = snapshot.GetSheet(sheetId);
        if (hidden == false && WorksheetOutline.Groups(sheet, axis).Any(g => g.Collapsed && g.First < first + count && g.Last >= first))
            throw new InvalidOperationException("该区域由分组折叠隐藏，请通过展开分组恢复可见性。");
        var rows = sheet.Rows.ToBuilder(); var columns = sheet.Columns.ToBuilder();
        // Showing an axis changes only existing metadata, without materializing every default row.
        var indexes = size is null && hidden != true
            ? (axis == SheetAxis.Row ? sheet.Rows.Keys : sheet.Columns.Keys).Where(i => i >= first && i < first + count)
            : Enumerable.Range(first, count);
        foreach (var i in indexes)
            if (axis == SheetAxis.Row) { var old = rows.GetValueOrDefault(i) ?? new(); rows[i] = old with { Height = size ?? old.Height, Hidden = hidden ?? old.Hidden }; }
            else { var old = columns.GetValueOrDefault(i) ?? new(); columns[i] = old with { Width = size ?? old.Width, Hidden = hidden ?? old.Hidden }; }
        if (rows.SequenceEqual(sheet.Rows) && columns.SequenceEqual(sheet.Columns)) return new(snapshot, [], false);
        var updated = new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells, rows, columns, sheet.MergedCells, sheet.Images,
            sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations, sheet.Data);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? updated : s)), [], true,
            CalculationUnchanged: size.HasValue && hidden is null);
    }
}

public sealed class DeleteSheetCommand(SheetId sheetId) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (snapshot.Sheets.Length == 1) throw new InvalidOperationException("至少保留一张工作表。");
        var sheet = snapshot.GetSheet(sheetId);
        var transformed = StructureEdits.Transform(snapshot, (node, _) => node switch
        {
            StructuredReferenceNode r when string.Equals(r.Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase) ||
                sheet.Data.Tables.Any(t => string.Equals(t.Name, r.Table, StringComparison.OrdinalIgnoreCase)) => new ValueNode(CellValue.FromError(CellErrorCode.Reference)),
            NameNode n when sheet.Data.Tables.Any(t => t.Name.Equals(n.Name, StringComparison.OrdinalIgnoreCase)) => new ValueNode(CellValue.FromError(CellErrorCode.Reference)),
            ReferenceNode r when string.Equals(r.Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase) => new ValueNode(CellValue.FromError(CellErrorCode.Reference)),
            NameNode n when string.Equals(n.Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase) => new ValueNode(CellValue.FromError(CellErrorCode.Reference)),
            _ => node
        });
        var remaining = transformed.WithDefinedNames(transformed.DefinedNames.Where(n => n.Scope != sheetId));
        var changes = StructureEdits.Draft(snapshot, transformed).Changes.Where(c => c.Address.SheetId != sheetId)
            .Concat(sheet.Cells.Select(c => new CellChange(new(sheetId, c.Key.Row, c.Key.Column), c.Value, null))).ToImmutableArray();
        return new(remaining.WithSheets(remaining.Sheets.Where(s => s.Id != sheetId)), changes, true);
    }
}
