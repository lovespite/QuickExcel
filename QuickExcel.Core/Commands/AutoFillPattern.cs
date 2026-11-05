using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Conservative lane inference: ambiguous, mixed, formula and rich-text seeds repeat unchanged.</summary>
internal static class AutoFillPattern
{
    internal static Func<int, CellValue>? Infer(CellContent[] seeds, WorkbookSnapshot snapshot)
    {
        if (seeds.Any(c => c.Kind != CellContentKind.Literal || c.RichText is not null)) return null;
        var first = seeds[0].LiteralValue;
        if (first.Kind == CellValueKind.Text) return TextPattern(seeds);
        if (first.Kind != CellValueKind.Number || seeds.Any(c => c.LiteralValue.Kind != CellValueKind.Number)) return null;
        var values = seeds.Select(c => c.LiteralValue.NumberValue).ToArray();
        var dates = seeds.Select(c => CellDisplayFormatter.IsDateFormat(snapshot.Styles[c.StyleId], c.LiteralValue.NumberValue)).ToArray();
        if (dates.Any(d => d) && !dates.All(d => d)) return null;
        var start = values[0];
        var step = values.Length == 1 ? 1 : values[1] - start;
        if (dates[0])
        {
            if (values.Length > 1 && CalendarMonth(start, snapshot.DateSystem) is { } firstMonth &&
                CalendarMonth(values[1], snapshot.DateSystem) is { } secondMonth)
            {
                var months = secondMonth - firstMonth;
                var kind = months % 12 == 0 ? CellSeriesKind.Years : CellSeriesKind.Months;
                var calendarStep = kind == CellSeriesKind.Years ? months / 12 : months;
                if (months != 0 && Matches(values, start, calendarStep, kind, snapshot.DateSystem))
                    return Numeric(start, calendarStep, kind, snapshot.DateSystem);
            }
            return Matches(values, start, step, CellSeriesKind.Days, snapshot.DateSystem)
                ? Numeric(start, step, CellSeriesKind.Days, snapshot.DateSystem) : null;
        }
        if (values.Length == 1 || !Matches(values, start, step, CellSeriesKind.Numbers, snapshot.DateSystem)) return null;
        return Numeric(start, step, CellSeriesKind.Numbers, snapshot.DateSystem);
    }

    private static Func<int, CellValue>? TextPattern(CellContent[] seeds)
    {
        var text = seeds[0].LiteralValue.TextValue;
        if (!FillSeriesCommand.TryParseTextSeries(text, out var prefix, out var number, out _, out _)) return null;
        long step = 1;
        if (seeds.Length > 1)
        {
            if (seeds[1].LiteralValue.Kind != CellValueKind.Text ||
                !FillSeriesCommand.TryParseTextSeries(seeds[1].LiteralValue.TextValue, out var nextPrefix, out var next, out _, out _) || nextPrefix != prefix) return null;
            step = next - number;
        }
        if (step is < -100_000 or > 100_000) return null;
        try
        {
            for (var i = 0; i < seeds.Length; i++)
                if (seeds[i].LiteralValue.Kind != CellValueKind.Text || seeds[i].LiteralValue.TextValue != FillSeriesCommand.TextSeriesValue(text, step, i)) return null;
        }
        catch (ArgumentException) { return null; }
        return index => CellValue.FromText(FillSeriesCommand.TextSeriesValue(text, step * Math.Sign(index), Math.Abs(index)));
    }

    private static Func<int, CellValue> Numeric(double start, double step, CellSeriesKind kind, ExcelDateSystem system) =>
        index => CellValue.FromNumber(FillSeriesCommand.SeriesValue(start, step * Math.Sign(index), Math.Abs(index), kind, system));

    private static bool Matches(double[] values, double start, double step, CellSeriesKind kind, ExcelDateSystem system)
    {
        try
        {
            for (var i = 0; i < values.Length; i++)
            {
                var expected = FillSeriesCommand.SeriesValue(start, step, i, kind, system);
                // Allow arithmetic round-off without treating visibly different large numbers as a trend.
                var tolerance = Math.Max(1, Math.Abs(step)) * 1e-10 + Math.Abs(expected) * 4e-16;
                if (Math.Abs(values[i] - expected) > tolerance) return false;
            }
            return true;
        }
        catch (ArgumentException) { return false; }
    }

    private static int? CalendarMonth(double serial, ExcelDateSystem system)
    {
        var date = new ExcelSerialDate(serial, system);
        if (date.IsFictitiousLeapDay) return 1900 * 12 + 1;
        return date.TryToDateTime(out var value) ? value.Year * 12 + value.Month - 1 : null;
    }
}
