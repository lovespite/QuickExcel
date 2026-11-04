using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public enum ChartSourceRole { Categories, Values, Name }
public sealed record ChartSourceRange(string Sheet, CellRectangle Range, ChartSourceRole Role, string Color);

public static class ChartSelection
{
    public static IReadOnlyList<ChartSourceRange> Sources(WorksheetChart chart)
    {
        var result = new List<ChartSourceRange>();
        void Add(string? formula, ChartSourceRole role, string color)
        {
            if (formula is null) return;
            var r = WorksheetChart.ParseReference(formula);
            if (result.Any(s => s.Sheet.Equals(r.Sheet, StringComparison.OrdinalIgnoreCase) && s.Range == new CellRectangle(r.Row, r.Column, r.EndRow, r.EndColumn) && s.Role == role)) return;
            result.Add(new(r.Sheet!, new(r.Row, r.Column, r.EndRow, r.EndColumn), role, color));
        }
        foreach (var series in chart.Series)
        {
            Add(series.Categories, ChartSourceRole.Categories, "FF9054C8");
            Add(series.Values, ChartSourceRole.Values, "FF377CE0");
            Add(series.NameReference, ChartSourceRole.Name, "FFE45F6B");
        }
        return result.OrderBy(s => s.Role).ToArray();
    }
}
