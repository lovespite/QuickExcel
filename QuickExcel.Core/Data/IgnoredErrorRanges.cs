using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Rectangle operations keep error-checking preferences attached during cell transfers.</summary>
internal static class IgnoredErrorRanges
{
    internal static ImmutableArray<WorksheetIgnoredError> Clip(IEnumerable<WorksheetIgnoredError> source, CellRectangle area) =>
        source.Where(e => ValidationSafety.Overlaps(e.Range, area))
            .Select(e => e with { Range = Intersection(e.Range, area) }).ToImmutableArray();

    internal static ImmutableArray<WorksheetIgnoredError> Replace(IEnumerable<WorksheetIgnoredError> source, CellRectangle area,
        IEnumerable<WorksheetIgnoredError>? replacement = null)
    {
        var result = source.SelectMany(e => Subtract(e.Range, area).Select(r => e with { Range = r }))
            .Concat(replacement ?? []).Distinct().Take(4097).ToImmutableArray();
        if (result.Length > 4096) throw new NotSupportedException(MessageText.Get("操作会使忽略错误提示的区域超过 4096 个，整批操作未应用。"));
        return result;
    }

    internal static CellRectangle Intersection(CellRectangle a, CellRectangle b) =>
        new(Math.Max(a.Top, b.Top), Math.Max(a.Left, b.Left), Math.Min(a.Bottom, b.Bottom), Math.Min(a.Right, b.Right));

    private static IEnumerable<CellRectangle> Subtract(CellRectangle source, CellRectangle removed)
    {
        if (!ValidationSafety.Overlaps(source, removed)) { yield return source; yield break; }
        var overlap = Intersection(source, removed);
        if (source.Top < overlap.Top) yield return source with { Bottom = overlap.Top - 1 };
        if (source.Bottom > overlap.Bottom) yield return source with { Top = overlap.Bottom + 1 };
        if (source.Left < overlap.Left) yield return new(overlap.Top, source.Left, overlap.Bottom, overlap.Left - 1);
        if (source.Right > overlap.Right) yield return new(overlap.Top, overlap.Right + 1, overlap.Bottom, source.Right);
    }
}
