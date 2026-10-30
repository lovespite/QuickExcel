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
            throw new InvalidOperationException("筛选列含未支持或依赖版面的显示格式，不能安全匹配显示值；请先改为明确的数字格式或使用数值条件。");

    public static ImmutableHashSet<int> Evaluate(WorkbookSnapshot snapshot, SheetId sheetId, bool allowSourceCaches = false,
        CancellationToken cancellationToken = default) => Evaluate(snapshot.GetSheet(sheetId), snapshot.Styles, snapshot.DateSystem,
            snapshot.Revision, allowSourceCaches, cancellationToken);

    public static ImmutableHashSet<int> Evaluate(WorksheetSnapshot sheet, IReadOnlyList<StyleDefinition> styles, ExcelDateSystem system,
        long revision, bool allowSourceCaches = false, CancellationToken cancellationToken = default)
    {
        if (sheet.Data.Filter is not { } filter || filter.Columns.Count == 0) return [];
        var predicates = filter.Columns.Select(p => (Column: p.Key, Match: Compile(p.Value))).ToArray();
        var hidden = ImmutableHashSet.CreateBuilder<int>();
        for (var row = filter.Range.Top + 1; row <= filter.Range.Bottom; row++)
        {
            cancellationToken.ThrowIfCancellationRequested(); var matches = true;
            foreach (var (column, match) in predicates)
            {
                var cell = sheet.GetCell(row, column);
                if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation?.Value is null || !allowSourceCaches &&
                    (cell.Calculation.InputRevision != revision || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error))))
                    throw new InvalidOperationException("筛选列包含待重算或缺少结果的公式，请重算后重新应用或保存。");
                matches &= match(cell?.Value ?? CellValue.Blank, styles[sheet.GetStyleId(row, column)], system);
            }
            if (!matches) hidden.Add(row);
        }
        return hidden.ToImmutable();
    }

    public static Func<CellValue, StyleDefinition, ExcelDateSystem, bool> Compile(AutoFilterColumn column)
    {
        if (column.Values is { } values)
        {
            var labels = values.ToHashSet(StringComparer.OrdinalIgnoreCase);
            return (value, style, system) => value.Kind == CellValueKind.Blank || value.Kind == CellValueKind.Text && value.TextValue.Length == 0
                ? column.IncludeBlank : labels.Contains(ValueLabel(value, style, system));
        }
        if (column.First is null) return (_, _, _) => true;
        var first = Condition(column.First); var second = column.Second is null ? null : Condition(column.Second);
        return (value, _, _) => second is null ? first(value) : column.UseAnd ? first(value) && second(value) : first(value) || second(value);
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
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId); var data = sheet.Data with { Filter = filter };
        if (sheet.Data.ContentEquals(data)) return new(snapshot, [], false);
        var next = sheet.WithData(data);
        if (filter is not null) _ = WorksheetFilterEngine.Evaluate(next, snapshot.Styles, snapshot.DateSystem, snapshot.Revision);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? next : s)), [], false) { CalculationUnchanged = true };
    }
}

public sealed class ClearSortStateCommand(SheetId sheetId) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Data.Sort is null) return new(snapshot, [], false);
        var next = sheet.WithData(sheet.Data with { Sort = null });
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? next : s)), [], false) { CalculationUnchanged = true };
    }
}
