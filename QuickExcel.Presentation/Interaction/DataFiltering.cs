using System.Collections.Immutable;
using System.Globalization;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public sealed record DataRegion(Guid Id, SheetId SheetId, GridRange Range, bool HasHeader = true)
{
    public int FirstDataRow => Range.Top + (HasHeader ? 1 : 0);
    public void Validate()
    {
        if (Range.Top < 0 || Range.Left < 0 || Range.Bottom >= WorkbookLimits.MaxRows || Range.Right >= WorkbookLimits.MaxColumns ||
            Range.Top > Range.Bottom || Range.Left > Range.Right)
            throw new ArgumentException($"筛选区域须是 {WorkbookLimits.MaxRows:N0} 行 × {WorkbookLimits.MaxColumns:N0} 列范围内、包含数据的矩形。");
    }
}

public enum FilterOperator { TextEquals, TextContains, TextStartsWith, TextEndsWith, NumberEquals, Greater, GreaterOrEqual, Less, LessOrEqual, DateOnOrAfter, DateOnOrBefore, ExcelEquals, ExcelNotEquals, ExcelLess, ExcelLessOrEqual, ExcelGreater, ExcelGreaterOrEqual }

public sealed record FilterTerm(FilterOperator Operator, string Operand)
{
    public void Validate()
    {
        if (!Enum.IsDefined(Operator) || Operand.Length > 1024) throw new ArgumentException("筛选条件无效或超过 1024 个字符。");
        if (Operator is >= FilterOperator.NumberEquals and <= FilterOperator.LessOrEqual &&
            (!double.TryParse(Operand, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)))
            throw new ArgumentException("数字条件须为有限数值（小数点使用 .）。");
        if (Operator is FilterOperator.DateOnOrAfter or FilterOperator.DateOnOrBefore &&
            !DateTime.TryParseExact(Operand, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            throw new ArgumentException("日期条件须为 yyyy-MM-dd。");
    }

    public bool Matches(CellValue value, StyleDefinition style, ExcelDateSystem dateSystem)
    {
        if (Operator >= FilterOperator.ExcelEquals) return WorksheetFilterEngine.Compile(new(first: new((AutoFilterComparison)(Operator - FilterOperator.ExcelEquals), Operand)))(value, style, dateSystem);
        if (Operator <= FilterOperator.TextEndsWith)
        {
            if (value.Kind != CellValueKind.Text) return false;
            return Operator switch
            {
                FilterOperator.TextEquals => string.Equals(value.TextValue, Operand, StringComparison.OrdinalIgnoreCase),
                FilterOperator.TextContains => value.TextValue.Contains(Operand, StringComparison.OrdinalIgnoreCase),
                FilterOperator.TextStartsWith => value.TextValue.StartsWith(Operand, StringComparison.OrdinalIgnoreCase),
                _ => value.TextValue.EndsWith(Operand, StringComparison.OrdinalIgnoreCase)
            };
        }
        if (value.Kind != CellValueKind.Number) return false;
        if (Operator is FilterOperator.DateOnOrAfter or FilterOperator.DateOnOrBefore)
        {
            if (!CellDisplayFormatter.IsDateFormat(style, value.NumberValue) ||
                !new ExcelSerialDate(value.NumberValue, dateSystem).TryToDateTime(out var date)) return false;
            var operand = DateTime.ParseExact(Operand, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            return Operator == FilterOperator.DateOnOrAfter ? date.Date >= operand : date.Date <= operand;
        }
        var number = double.Parse(Operand, CultureInfo.InvariantCulture);
        return Operator switch
        {
            FilterOperator.NumberEquals => value.NumberValue == number,
            FilterOperator.Greater => value.NumberValue > number,
            FilterOperator.GreaterOrEqual => value.NumberValue >= number,
            FilterOperator.Less => value.NumberValue < number,
            _ => value.NumberValue <= number
        };
    }
}

/// <summary>Exact typed value membership OR two typed predicates. Search is a UI concern and never edits this set.</summary>
public sealed record ColumnFilter(ImmutableHashSet<CellValue>? Values = null, FilterTerm? First = null, FilterTerm? Second = null, bool UseOr = false, AutoFilterColumn? Stored = null)
{
    public void Validate()
    {
        if (Values is not null && (First is not null || Second is not null) || Second is not null && First is null)
            throw new ArgumentException("候选值筛选和条件筛选不能混用。");
        First?.Validate(); Second?.Validate();
    }
    public bool Matches(CellValue value, StyleDefinition style, ExcelDateSystem system) => Stored is not null ? WorksheetFilterEngine.Compile(Stored)(value, style, system) : Values is not null ? Values.Contains(value)
        : First is null || (Second is null ? First.Matches(value, style, system) : UseOr
            ? First.Matches(value, style, system) || Second.Matches(value, style, system)
            : First.Matches(value, style, system) && Second.Matches(value, style, system));
}

public sealed record FilterCandidate(CellValue Key, string Label, int Count, bool Selected);

public static class DataFilterEngine
{
    public const int MaxCandidateCharacters = 8 * 1024 * 1024;

    private static (CellValue Value, StyleDefinition Style) Read(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column)
    {
        var cell = sheet.GetCell(row, column);
        if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation?.Value is null ||
            cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.InputRevision != snapshot.Revision))
            throw new InvalidOperationException("筛选列包含待重算公式，请完成重算后重新应用。");
        return (cell?.Value ?? CellValue.Blank, snapshot.Styles[sheet.GetStyleId(row, column)]);
    }

    private static WorksheetSnapshot Validate(WorkbookSnapshot snapshot, DataRegion region, IReadOnlyDictionary<int, ColumnFilter> filters)
    {
        region.Validate();
        foreach (var (column, filter) in filters)
        {
            if (column < region.Range.Left || column > region.Range.Right) throw new ArgumentException("筛选列不在数据区域中。");
            filter.Validate();
        }
        return snapshot.GetSheet(region.SheetId);
    }

    private static bool Matches(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, IReadOnlyDictionary<int, ColumnFilter> filters, int? except = null)
    {
        var matches = true;
        foreach (var (column, filter) in filters)
        {
            if (column == except) continue;
            var cell = Read(snapshot, sheet, row, column);
            // Do not short circuit: stale results in any dependent column must reject the whole result.
            matches &= filter.Matches(cell.Value, cell.Style, snapshot.DateSystem);
        }
        return matches;
    }

    public static ImmutableHashSet<int> Evaluate(WorkbookSnapshot snapshot, DataRegion region,
        IReadOnlyDictionary<int, ColumnFilter> filters, CancellationToken cancellationToken = default)
    {
        var sheet = Validate(snapshot, region, filters);
        var hidden = ImmutableHashSet.CreateBuilder<int>();
        for (var row = region.FirstDataRow; row <= region.Range.Bottom; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!Matches(snapshot, sheet, row, filters)) hidden.Add(row);
        }
        return hidden.ToImmutable();
    }

    public static IReadOnlyList<FilterCandidate> Candidates(WorkbookSnapshot snapshot, DataRegion region,
        IReadOnlyDictionary<int, ColumnFilter> filters, int column, CancellationToken cancellationToken = default)
    {
        var sheet = Validate(snapshot, region, filters);
        if (column < region.Range.Left || column > region.Range.Right) throw new ArgumentException("候选列不在区域内。");
        var selected = filters.GetValueOrDefault(column)?.Values;
        var values = new Dictionary<CellValue, (string Label, int Count)>();
        var characters = 0;
        void Add(CellValue value, string label, int count)
        {
            if (values.TryGetValue(value, out var existing)) { values[value] = (existing.Label, existing.Count + count); return; }
            characters += label.Length;
            if (characters > MaxCandidateCharacters) throw new InvalidOperationException("候选文本超过 8 Mi 字符，请缩小筛选区域；没有截断或应用筛选。");
            values.Add(value, (label, count));
        }
        for (var row = region.FirstDataRow; row <= region.Range.Bottom; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = Read(snapshot, sheet, row, column);
            if (Matches(snapshot, sheet, row, filters, column))
                Add(cell.Value, Label(cell.Value, cell.Style, snapshot.DateSystem), 1);
        }
        if (selected is not null)
            foreach (var value in selected) Add(value, Label(value, new(), snapshot.DateSystem), 0);
        return values.Select(pair => new FilterCandidate(pair.Key, pair.Value.Label, pair.Value.Count, selected is null || selected.Contains(pair.Key)))
            .OrderBy(candidate => candidate.Key.Kind == CellValueKind.Blank).ThenBy(candidate => candidate.Label, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static string Label(CellValue value, StyleDefinition style, ExcelDateSystem system) => value.Kind switch
    {
        CellValueKind.Blank => "(空白)",
        CellValueKind.Text => value.TextValue.Length == 0 ? "(空字符串)" : value.TextValue + "  [文本]",
        CellValueKind.Number => CellDisplayFormatter.Format(value, style, system) + "  [数字 " + value.NumberValue.ToString("R", CultureInfo.InvariantCulture) + "]",
        CellValueKind.Boolean => CellInput.DisplayText(value) + "  [布尔]",
        _ => CellInput.DisplayText(value) + "  [错误]"
    };
}
