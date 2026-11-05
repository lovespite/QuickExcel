using System.Collections.Immutable;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public static class WorksheetFilterEngine
{
    public static bool CanUseSourceCaches(WorkbookSnapshot snapshot) => snapshot.FileState is { } source &&
        (source.FormulaReadOnly || snapshot.Revision == source.OriginalModel.Revision);
    public static string ValueLabel(CellValue value, StyleDefinition style, ExcelDateSystem system) =>
        CellDisplayFormatter.SupportsFilterLabel(value, style) ? CellDisplayFormatter.Format(value, style, system) :
            throw new InvalidOperationException(MessageText.Get("筛选列含未支持或依赖版面的显示格式，不能安全匹配显示值；请先改为明确的数字格式或使用数值条件。"));

    public static ImmutableHashSet<int> Evaluate(WorkbookSnapshot snapshot, SheetId sheetId, bool allowSourceCaches = false,
        CancellationToken cancellationToken = default) => Evaluate(snapshot.GetSheet(sheetId), snapshot.Styles, snapshot.DateSystem,
            snapshot.Revision, allowSourceCaches, cancellationToken);

    public static ImmutableHashSet<int> Evaluate(WorksheetSnapshot sheet, IReadOnlyList<StyleDefinition> styles, ExcelDateSystem system,
        long revision, bool allowSourceCaches = false, CancellationToken cancellationToken = default)
    {
        var hidden = ImmutableHashSet.CreateBuilder<int>();
        if (sheet.Data.AllFilters.Sum(f => (long)(f.Range.Bottom - f.Range.Top) * f.Columns.Count) > 2_000_000)
            throw new InvalidOperationException(MessageText.Get("工作表筛选条件遍历超过 2,000,000 格。"));
        foreach (var filter in sheet.Data.AllFilters)
        {
        if (filter.Columns.Count == 0) continue;
        RequireColorScope(sheet, filter);
        CellValue ReadValue(int row, int column)
        {
            cancellationToken.ThrowIfCancellationRequested(); var cell = sheet.GetCell(row, column);
            if ((cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null) && (cell.Calculation?.Value is null || !allowSourceCaches &&
                (cell.Calculation.InputRevision != revision || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error))))
                throw new InvalidOperationException(MessageText.Get("筛选列包含待重算或缺少结果的公式，请重算后重新应用或保存。"));
            return cell?.Value ?? CellValue.Blank;
        }
        var predicates = filter.Columns.Select(p => (Column: p.Key, Match: Compile(p.Value,
            p.Value.Top is null ? null : Enumerable.Range(filter.Range.Top + 1, filter.Range.Rows - 1).Select(row => ReadValue(row, p.Key))))).ToArray();
        for (var row = filter.Range.Top + 1; row <= filter.Range.Bottom; row++)
        {
            cancellationToken.ThrowIfCancellationRequested(); var matches = true;
            foreach (var (column, match) in predicates)
            {
                matches &= match(ReadValue(row, column), TablePresentation.NumberStyle(sheet, styles, row, column), system);
            }
            if (!matches) hidden.Add(row);
        }
        }
        return hidden.ToImmutable();
    }

    public static Func<CellValue, StyleDefinition, ExcelDateSystem, bool> Compile(AutoFilterColumn column, IEnumerable<CellValue>? population = null)
    {
        if (column.Top is { } top)
        {
            if (population is null) throw new InvalidOperationException(MessageText.Get("前后项筛选需要完整筛选列上下文。"));
            var populationValues = population.ToArray();
            if (populationValues.Any(v => v.Kind == CellValueKind.Error)) throw new InvalidOperationException(MessageText.Get("前后项筛选列包含错误值，请先修正错误。"));
            var numbers = populationValues.Where(v => v.Kind == CellValueKind.Number).Select(v => v.NumberValue).Order().ToArray();
            if (numbers.Length == 0) throw new InvalidOperationException(MessageText.Get("前后项筛选列没有可排名的数值。"));
            var count = Math.Min(numbers.Length, top.Percent ? Math.Max(1, numbers.Length * top.Count / 100) : top.Count);
            var threshold = numbers[top.Bottom ? count - 1 : numbers.Length - count];
            return (value, _, _) => value.Kind == CellValueKind.Number && (top.Bottom ? value.NumberValue <= threshold : value.NumberValue >= threshold);
        }
        if (column.Color is { } color)
            return (_, style, _) => string.Equals(color.Font ? style.FontColor ?? "FF000000" : style.FillColor, color.Argb, StringComparison.OrdinalIgnoreCase);
        if (column.Values is { } values)
        {
            var labels = values.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var years = column.DateGroups.Where(g => g.Month is null).Select(g => g.Year).ToHashSet();
            var months = column.DateGroups.Where(g => g.Month is not null && g.Day is null).Select(g => (g.Year, g.Month!.Value)).ToHashSet();
            var days = column.DateGroups.Where(g => g.Day is not null).Select(g => (g.Year, g.Month!.Value, g.Day!.Value)).ToHashSet();
            bool DateMatches(CellValue value, StyleDefinition style, ExcelDateSystem system)
            {
                if (column.DateGroups.IsEmpty || value.Kind != CellValueKind.Number || value.NumberValue < 0 || !CellDisplayFormatter.IsDateFormat(style, value.NumberValue)) return false;
                var serial = new ExcelSerialDate(Math.Floor(value.NumberValue), system);
                int year, month, day;
                if (serial.IsFictitiousLeapDay) { year = 1900; month = 2; day = 29; }
                else if (serial.TryToDateTime(out var date)) { year = date.Year; month = date.Month; day = date.Day; }
                else return false;
                return years.Contains(year) || months.Contains((year, month)) || days.Contains((year, month, day));
            }
            return (value, style, system) => value.Kind == CellValueKind.Blank || value.Kind == CellValueKind.Text && value.TextValue.Length == 0
                ? column.IncludeBlank : DateMatches(value, style, system) || labels.Count > 0 && labels.Contains(ValueLabel(value, style, system));
        }
        if (column.First is null) return (_, _, _) => true;
        var first = Condition(column.First); var second = column.Second is null ? null : Condition(column.Second);
        return (value, _, _) => second is null ? first(value) : column.UseAnd ? first(value) && second(value) : first(value) || second(value);
    }

    public static void RequireColorScope(WorksheetSnapshot sheet, WorksheetAutoFilter filter)
    {
        foreach (var (column, rule) in filter.Columns)
        {
            if (rule.Color is null) continue;
            var range = filter.Range with { Left = column, Right = column };
            if (sheet.Data.Tables.Any(t => ValidationSafety.Overlaps(t.Range, range)) ||
                sheet.Data.ConditionalFormats.Any(f => f.Ranges.Any(r => ValidationSafety.Overlaps(r, range))))
                throw new InvalidOperationException(MessageText.Get("颜色筛选采用基础单元格颜色，暂不支持含 Table 样式或条件格式的筛选列。"));
        }
    }

    private static Func<CellValue, bool> Condition(AutoFilterCondition condition)
    {
        var op = condition.Comparison; var operand = condition.Operand;
        bool Compare(int c) => op switch { AutoFilterComparison.Equal => c == 0, AutoFilterComparison.NotEqual => c != 0,
            AutoFilterComparison.LessThan => c < 0, AutoFilterComparison.LessThanOrEqual => c <= 0,
            AutoFilterComparison.GreaterThan => c > 0, _ => c >= 0 };
        if (double.TryParse(operand, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
            return value => value.Kind == CellValueKind.Number ? Compare(value.NumberValue.CompareTo(number)) : op == AutoFilterComparison.NotEqual;
        if (bool.TryParse(operand, out var boolean))
            return value => value.Kind == CellValueKind.Boolean ? Compare(value.BooleanValue.CompareTo(boolean)) : op == AutoFilterComparison.NotEqual;
        if (Enum.GetValues<CellErrorCode>().FirstOrDefault(code => CellInput.DisplayText(CellValue.FromError(code)) == operand) is var error &&
            CellInput.DisplayText(CellValue.FromError(error)) == operand)
            return value => value.Kind == CellValueKind.Error ? Compare(value.ErrorCode.CompareTo(error)) : op == AutoFilterComparison.NotEqual;
        if (operand == " " && op is AutoFilterComparison.Equal or AutoFilterComparison.NotEqual)
            return value => (op == AutoFilterComparison.NotEqual) ^ (value.Kind == CellValueKind.Blank || value.Kind == CellValueKind.Text && value.TextValue.Length == 0);
        var pattern = new StringBuilder("\\A");
        for (var i = 0; i < operand.Length; i++)
        {
            var c = operand[i];
            if (c == '~' && i + 1 < operand.Length && operand[i + 1] is '*' or '?' or '~') pattern.Append(Regex.Escape(operand[++i].ToString()));
            else pattern.Append(c == '*' ? ".*" : c == '?' ? "." : Regex.Escape(c.ToString()));
        }
        pattern.Append("\\z");
        var regex = new Regex(pattern.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline | RegexOptions.NonBacktracking);
        return value =>
        {
            if (value.Kind == CellValueKind.Blank) return operand.Length == 0 ? Compare(0) : op == AutoFilterComparison.NotEqual;
            if (value.Kind != CellValueKind.Text) return op == AutoFilterComparison.NotEqual;
            return op is AutoFilterComparison.Equal or AutoFilterComparison.NotEqual ? (op == AutoFilterComparison.NotEqual) ^ regex.IsMatch(value.TextValue)
                : Compare(StringComparer.OrdinalIgnoreCase.Compare(value.TextValue, operand));
        };
    }

    public static string EscapeWildcard(string text) => text.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");
}

public sealed class SetAutoFilterCommand(SheetId sheetId, WorksheetAutoFilter? filter) : DocumentCommand
{
    private bool _valuesValidated;
    internal SetAutoFilterCommand WithValidatedValues() => new(sheetId, filter) { _valuesValidated = true };
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorksheetOutline.RequireEditable(snapshot);
        var sheet = snapshot.GetSheet(sheetId); var data = sheet.Data with { Filter = filter };
        if (sheet.Data.ContentEquals(data)) return new(snapshot, [], false);
        var next = sheet.WithData(data);
        if (filter is not null && !_valuesValidated) _ = WorksheetFilterEngine.Evaluate(next, snapshot.Styles, snapshot.DateSystem, snapshot.Revision);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? next : s)), [], false) { CalculationUnchanged = true };
    }
}

public sealed class ClearSortStateCommand(SheetId sheetId, int? tableId = null) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (tableId is { } id)
        {
            var table = sheet.Data.Tables.Single(t => t.Id == id); if (table.Sort is null) return new(snapshot, [], false);
            var updated = sheet.WithData(sheet.Data with { Tables = sheet.Data.Tables.Replace(table, table.WithSort(null)) });
            return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? updated : s)), [], false) { CalculationUnchanged = true };
        }
        if (sheet.Data.Sort is null) return new(snapshot, [], false);
        var next = sheet.WithData(sheet.Data with { Sort = null });
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? next : s)), [], false) { CalculationUnchanged = true };
    }
}
