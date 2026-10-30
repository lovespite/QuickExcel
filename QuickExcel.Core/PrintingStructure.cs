using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

internal static class PrintingStructure
{
    internal static WorksheetPrinting Axis(WorksheetPrinting printing, SheetAxis axis,
        Func<int, int?> position, Func<int, int, (int Start, int End)?> interval)
    {
        ImmutableArray<PrintRange> Ranges(ImmutableArray<PrintRange> ranges)
        {
            var result = ImmutableArray.CreateBuilder<PrintRange>();
            foreach (var range in ranges)
            {
                if (axis == SheetAxis.Row && range.Kind == PrintRangeKind.Columns || axis == SheetAxis.Column && range.Kind == PrintRangeKind.Rows)
                { result.Add(range); continue; }
                var mapped = interval(axis == SheetAxis.Row ? range.Top : range.Left, axis == SheetAxis.Row ? range.Bottom : range.Right);
                if (mapped is not { } next) continue;
                result.Add(axis == SheetAxis.Row ? range with { Top = next.Start, Bottom = next.End } : range with { Left = next.Start, Right = next.End });
            }
            return result.ToImmutable();
        }
        ImmutableArray<PrintPageBreak> Breaks(ImmutableArray<PrintPageBreak> breaks, SheetAxis breakAxis)
        {
            var result = ImmutableArray.CreateBuilder<PrintPageBreak>();
            foreach (var item in breaks)
            {
                if (axis == breakAxis)
                {
                    if (position(item.Id) is { } id && id > 0) result.Add(item with { Id = id });
                }
                else
                {
                    var full = axis == SheetAxis.Row ? PrintRange.MaxRows - 1 : PrintRange.MaxColumns - 1;
                    if (item.Minimum == 0 && item.Maximum == full) result.Add(item);
                    else if (interval(item.Minimum, item.Maximum) is { } span) result.Add(item with { Minimum = span.Start, Maximum = span.End });
                }
            }
            return result.OrderBy(b => b.Id).ToImmutableArray();
        }
        return printing with { Areas = Ranges(printing.Areas), Titles = Ranges(printing.Titles),
            RowBreaks = Breaks(printing.RowBreaks, SheetAxis.Row), ColumnBreaks = Breaks(printing.ColumnBreaks, SheetAxis.Column) };
    }
}
