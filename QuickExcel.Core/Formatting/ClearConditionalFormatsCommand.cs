using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Remove rules from selected rectangles without changing their meaning outside the selection.</summary>
public sealed class ClearConditionalFormatsCommand : DocumentCommand
{
    private readonly SheetId _sheet;
    private readonly ImmutableArray<CellRectangle>? _ranges;
    public ClearConditionalFormatsCommand(SheetId sheet, IEnumerable<CellRectangle>? ranges = null)
    {
        _sheet = sheet;
        if (ranges is null) return;
        var selection = ranges.Take(1025).ToImmutableArray();
        if (selection.Length is < 1 or > 1024) throw new ArgumentException("请选择 1 到 1024 个区域。");
        foreach (var range in selection)
        {
            WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
            if (range.Top > range.Bottom || range.Left > range.Right) throw new ArgumentException("选区无效。");
        }
        _ranges = selection;
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var rules = new List<ConditionalFormatRule>();
        if (_ranges is { } selection)
            foreach (var rule in snapshot.GetSheet(_sheet).Data.ConditionalFormats)
            {
                var remaining = rule.Ranges;
                foreach (var cut in selection)
                {
                    remaining = remaining.SelectMany(r => Subtract(r, cut)).Take(1025).ToImmutableArray();
                    if (remaining.Length > 1024) throw new NotSupportedException("清除后规则区域超过 1024，整批操作未应用。");
                }
                if (remaining.IsEmpty) continue;
                if (remaining.SequenceEqual(rule.Ranges)) { rules.Add(rule); continue; }
                var anchor = new CellPosition(remaining[0].Top, remaining[0].Left);
                rules.Add(rule.MapFormulas(f => FormulaCopy.Translate(f, rule.Anchor, anchor)).WithRanges(remaining));
            }
        // Reuse unknown-content preservation and binding validation; the session enforces protection.
        return new SetConditionalFormatsCommand(_sheet, rules).Prepare(snapshot);
    }

    private static IEnumerable<CellRectangle> Subtract(CellRectangle source, CellRectangle cut)
    {
        var top = Math.Max(source.Top, cut.Top); var bottom = Math.Min(source.Bottom, cut.Bottom);
        var left = Math.Max(source.Left, cut.Left); var right = Math.Min(source.Right, cut.Right);
        if (top > bottom || left > right) { yield return source; yield break; }
        if (source.Top < top) yield return new(source.Top, source.Left, top - 1, source.Right);
        if (bottom < source.Bottom) yield return new(bottom + 1, source.Left, source.Bottom, source.Right);
        if (source.Left < left) yield return new(top, source.Left, bottom, left - 1);
        if (right < source.Right) yield return new(top, right + 1, bottom, source.Right);
    }
}
