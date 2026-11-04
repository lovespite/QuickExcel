using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Replace the ordered worksheet rule collection in one undoable transaction.</summary>
public sealed class SetConditionalFormatsCommand(SheetId sheetId, IEnumerable<ConditionalFormatRule> rules) : DocumentCommand
{
    private readonly ImmutableArray<ConditionalFormatRule> _rules = rules.Take(1025).ToImmutableArray();
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase)) == true)
            throw new NotSupportedException("工作簿含尚未建模的条件格式；为保留完整优先级，规则编辑未应用。");
        var sheet = snapshot.GetSheet(sheetId);
        foreach (var rule in _rules) rule.ValidateBinding(snapshot, sheetId);
        var data = sheet.Data with { ConditionalFormats = _rules };
        if (sheet.Data.ContentEquals(data)) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(data) : s)), [], false) { CalculationUnchanged = true };
    }
}

internal static class ConditionalStructure
{
    internal static WorksheetData Normalize(WorksheetSnapshot before, WorksheetData after, SheetAxis axis,
        Func<int, int?> position, Func<FormulaNode, SheetId, CellPosition?, FormulaNode> transform)
    {
        return after with { ConditionalFormats = after.ConditionalFormats.Select(rule =>
        {
            var old = before.Data.ConditionalFormats.Single(r => r.Priority == rule.Priority);
            // The old top-left cell may have been deleted. Use the first surviving cell's
            // relative formula before applying the reference transformation.
            CellPosition? source = null;
            foreach (var range in old.Ranges)
            {
                var first = axis == SheetAxis.Row ? range.Top : range.Left;
                var last = axis == SheetAxis.Row ? range.Bottom : range.Right;
                if (position(first) is null)
                {
                    if (position(last) is null) continue;
                    var lo = first; var hi = last;
                    while (lo < hi) { var mid = lo + (hi - lo) / 2; if (position(mid) is null) lo = mid + 1; else hi = mid; }
                    first = lo;
                }
                source = axis == SheetAxis.Row ? new(first, range.Left) : new(range.Top, first); break;
            }
            if (source is not { } anchor) throw new InvalidOperationException("条件格式缺少存续的公式锚点。");
            var mapped = axis == SheetAxis.Row ? anchor with { Row = position(anchor.Row)!.Value } : anchor with { Column = position(anchor.Column)!.Value };
            return old.MapFormulas(formula =>
            {
                var copied = FormulaCopy.Translate(formula, old.Anchor, anchor);
                var expression = FormulaSyntax.Map(FormulaSyntax.Parse(copied), n => transform(n, before.Id, anchor));
                return FormulaCopy.Translate("=" + FormulaSyntax.Print(expression), mapped, rule.Anchor);
            }).WithRanges(rule.Ranges);
        }).ToImmutableArray() };
    }
}
