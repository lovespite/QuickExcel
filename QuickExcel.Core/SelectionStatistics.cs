using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record SelectionStatisticsResult(bool Complete, int Count, int Numbers, int Errors, int OldCaches, int MissingResults,
    double? Sum, double? Average, int VisitedCells, string? Limitation = null);

/// <summary>Sparse read-only statistics. Hidden rows/columns are excluded; overlapping selections count each stored cell once.</summary>
public static class SelectionStatistics
{
    public const int MaxRanges = 64;
    public const int MaxComparisons = 2_000_000;
    public static SelectionStatisticsResult Compute(WorkbookSnapshot snapshot, SheetId sheetId, IEnumerable<CellRectangle> selection,
        Func<int, bool>? filteredRow = null, CancellationToken cancellationToken = default, int comparisonBudget = MaxComparisons)
    {
        if (comparisonBudget is < 1 or > MaxComparisons) throw new ArgumentOutOfRangeException(nameof(comparisonBudget));
        var ranges = selection.Take(MaxRanges + 1).ToArray();
        SelectionStatisticsResult Incomplete(int visited, string reason) => new(false, 0, 0, 0, 0, 0, null, null, visited, reason);
        if (ranges.Length > MaxRanges) return Incomplete(0, "选区超过 64 个，请缩小选区。");
        foreach (var range in ranges)
        {
            WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
            if (range.Top > range.Bottom || range.Left > range.Right) throw new ArgumentException("选区端点无效。");
        }
        var sheet = snapshot.GetSheet(sheetId);
        var count = 0; var numbers = 0; var errors = 0; var old = 0; var missing = 0; var visited = 0; var comparisons = 0;
        double sum = 0, compensation = 0, mean = 0;
        foreach (var (position, cell) in sheet.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested(); visited++;
            var selected = false;
            foreach (var range in ranges)
            {
                if (++comparisons > comparisonBudget) return Incomplete(visited, "选区统计达到遍历上限，请缩小选区。");
                if (range.Contains(position.Row, position.Column)) { selected = true; break; }
            }
            if (!selected || sheet.Rows.GetValueOrDefault(position.Row)?.Hidden == true || sheet.Columns.GetValueOrDefault(position.Column)?.Hidden == true || filteredRow?.Invoke(position.Row) == true) continue;
            if (cell.Content.Kind == CellContentKind.Formula)
            {
                if (cell.Calculation?.Value is null) { missing++; continue; }
                if (cell.Calculation.InputRevision != snapshot.Revision || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)) old++;
            }
            var value = cell.Value;
            if (value.Kind == CellValueKind.Blank) continue;
            count++;
            if (value.Kind == CellValueKind.Error) { errors++; continue; }
            if (value.Kind != CellValueKind.Number) continue;
            numbers++; var number = value.NumberValue;
            var next = sum + number;
            compensation += Math.Abs(sum) >= Math.Abs(number) ? (sum - next) + number : (number - next) + sum;
            sum = next;
            mean = mean * ((double)(numbers - 1) / numbers) + number / numbers;
        }
        var total = sum + compensation;
        return new(true, count, numbers, errors, old, missing, numbers == 0 ? 0 : double.IsFinite(total) ? total : null,
            numbers == 0 ? null : double.IsFinite(total) ? total / numbers : double.IsFinite(mean) ? mean : null, visited);
    }
}
