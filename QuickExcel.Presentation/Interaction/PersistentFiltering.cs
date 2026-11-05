using System.Collections.Immutable;
using System.Globalization;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public static class PersistentFiltering
{
    public static AutoFilterColumn Convert(ColumnFilter filter, ExcelDateSystem system)
    {
        if (filter.Stored is { } stored) return stored;
        filter.Validate();
        if (filter.Values is { } values) return new(values.Where(v => v.Kind != CellValueKind.Blank && !(v.Kind == CellValueKind.Text && v.TextValue.Length == 0))
            .Select(v => v.Kind == CellValueKind.Text ? v.TextValue : WorksheetFilterEngine.ValueLabel(v, new(), system)).Order(StringComparer.OrdinalIgnoreCase),
            values.Any(v => v.Kind == CellValueKind.Blank || v.Kind == CellValueKind.Text && v.TextValue.Length == 0));
        AutoFilterCondition Condition(FilterTerm term)
        {
            if (term.Operator >= FilterOperator.ExcelEquals) return new((AutoFilterComparison)(term.Operator - FilterOperator.ExcelEquals), term.Operand);
            var text = term.Operator == FilterOperator.TextEquals ? term.Operand.Replace("*", "~*").Replace("?", "~?") : WorksheetFilterEngine.EscapeWildcard(term.Operand);
            if (term.Operator <= FilterOperator.TextEndsWith)
                return new(AutoFilterComparison.Equal, term.Operator switch { FilterOperator.TextContains => "*" + text + "*", FilterOperator.TextStartsWith => text + "*", FilterOperator.TextEndsWith => "*" + text, _ => text });
            if (term.Operator is FilterOperator.DateOnOrAfter or FilterOperator.DateOnOrBefore)
            {
                var date = DateTime.ParseExact(term.Operand, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                var serial = ExcelSerialDate.FromDateTime(date, system).Serial;
                return new(term.Operator == FilterOperator.DateOnOrAfter ? AutoFilterComparison.GreaterThanOrEqual : AutoFilterComparison.LessThan,
                    (term.Operator == FilterOperator.DateOnOrAfter ? serial : serial + 1).ToString("R", CultureInfo.InvariantCulture));
            }
            return new(term.Operator switch { FilterOperator.NumberEquals => AutoFilterComparison.Equal, FilterOperator.Greater => AutoFilterComparison.GreaterThan,
                FilterOperator.GreaterOrEqual => AutoFilterComparison.GreaterThanOrEqual, FilterOperator.Less => AutoFilterComparison.LessThan, _ => AutoFilterComparison.LessThanOrEqual }, term.Operand);
        }
        return new(first: filter.First is null ? null : Condition(filter.First), second: filter.Second is null ? null : Condition(filter.Second), useAnd: filter.Second is not null && !filter.UseOr);
    }

    public static ColumnFilter Restore(AutoFilterColumn column) => new(First: column.First is null ? null : new((FilterOperator)((int)FilterOperator.ExcelEquals + (int)column.First.Comparison), column.First.Operand),
        Second: column.Second is null ? null : new((FilterOperator)((int)FilterOperator.ExcelEquals + (int)column.Second.Comparison), column.Second.Operand), UseOr: !column.UseAnd, Stored: column);

    public static IReadOnlyList<FilterCandidate> Candidates(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int column, int? tableId = null)
    {
        var table = tableId is { } id ? sheet.Data.Tables.Single(t => t.Id == id) : null;
        var filter = (table is null ? sheet.Data.Filter : table.Filter) ?? throw new InvalidOperationException(UiText.Get("筛选区域已失效。"));
        var remaining = new WorksheetAutoFilter(filter.Range, filter.Columns.Remove(column));
        var other = sheet.WithData(table is null ? sheet.Data with { Filter = remaining } : sheet.Data with { Tables = sheet.Data.Tables.Replace(table, table.WithFilter(remaining)) });
        var hidden = WorksheetFilterEngine.Evaluate(other, snapshot.Styles, snapshot.DateSystem, snapshot.Revision);
        var selected = filter.Columns.GetValueOrDefault(column); var match = selected is null ? null : WorksheetFilterEngine.Compile(selected);
        var values = new Dictionary<CellValue, (string Label, int Count, bool Selected)>(); long characters = 0;
        void Add(CellValue key, string label, bool on, int count)
        {
            if (values.TryGetValue(key, out var old)) { values[key] = (old.Label, old.Count + count, old.Selected || on); return; }
            characters += label.Length;
            if (characters > DataFilterEngine.MaxCandidateCharacters) throw new InvalidOperationException(UiText.Get("筛选候选文本过大，请缩小区域。"));
            values.Add(key, (label, count, on));
        }
        for (var row = filter.Range.Top + 1; row <= filter.Range.Bottom; row++)
        {
            var cell = sheet.GetCell(row, column);
            if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation?.Value is null || cell.Calculation.InputRevision != snapshot.Revision || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)))
                throw new InvalidOperationException(UiText.Get("候选列含待重算公式，请先重算。"));
            if (hidden.Contains(row)) continue;
            var value = cell?.Value ?? CellValue.Blank; var style = TablePresentation.NumberStyle(sheet, snapshot.Styles, row, column);
            var blank = value.Kind == CellValueKind.Blank || value.Kind == CellValueKind.Text && value.TextValue.Length == 0;
            var label = blank ? UiText.Get("(空白)") : WorksheetFilterEngine.ValueLabel(value, style, snapshot.DateSystem);
            Add(blank ? CellValue.Blank : CellValue.FromText(label), label, match?.Invoke(value, style, snapshot.DateSystem) ?? true, 1);
        }
        if (selected?.Values is { } oldValues) foreach (var label in oldValues) Add(CellValue.FromText(label), label, true, 0);
        if (selected?.IncludeBlank == true) Add(CellValue.Blank, UiText.Get("(空白)"), true, 0);
        return values.Select(p => new FilterCandidate(p.Key, p.Value.Label, p.Value.Count, p.Value.Selected)).OrderBy(c => c.Key.Kind == CellValueKind.Blank).ThenBy(c => c.Label, StringComparer.OrdinalIgnoreCase).ToArray();
    }
}
