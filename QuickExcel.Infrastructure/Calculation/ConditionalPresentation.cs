using System.Globalization;
using System.Text;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

/// <summary>Viewport-demanded formatting projection. A new immutable snapshot invalidates the renderer's cache.</summary>
public static partial class ConditionalPresentation
{
    [ThreadStatic] private static Cache? _cache;
    private static Cache For(WorkbookSnapshot snapshot) => _cache?.Snapshot == snapshot ? _cache : _cache = new(snapshot);
    public static StyleDefinition GetStyle(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column)
    {
        if (sheet.Data.ConditionalFormats.IsEmpty) return TablePresentation.GetStyle(snapshot, sheet, row, column);
        return For(snapshot).Style(sheet, row, column);
    }
    public static IReadOnlyList<string> Diagnostics(WorkbookSnapshot snapshot) => _cache?.Snapshot == snapshot ? _cache.Errors.ToArray() : [];
    public static ConditionalCellVisual? GetVisual(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column) =>
        sheet.Data.ConditionalFormats.IsEmpty ? null : For(snapshot).Visual(sheet, row, column);
    private sealed partial class Cache(WorkbookSnapshot snapshot)
    {
        public WorkbookSnapshot Snapshot { get; } = snapshot;
        public HashSet<string> Errors { get; } = [];
        private readonly WorkbookValidation.Evaluation _evaluation = new(snapshot, useFreshCaches: true);
        private readonly Dictionary<CellAddress, Projection> _raw = [];
        private readonly Dictionary<CellAddress, StyleDefinition> _styles = [];
        private readonly Dictionary<(SheetId, int), Distribution> _counts = [];
        private long _countedCells;
        private sealed record Distribution(Dictionary<string, long> Keys, Dictionary<CellValue, long> Texts)
        {
            public Dictionary<CellValue, long> Wildcards { get; } = [];
        }
        private sealed record Projection(StyleDefinition Style, long[] Ranks, ConditionalCellVisual? Visual);
        public ConditionalCellVisual? Visual(WorksheetSnapshot sheet, int row, int column)
        { _evaluation.ResetBudget(); return Raw(sheet, row, column).Visual; }
        public StyleDefinition Style(WorksheetSnapshot sheet, int row, int column)
        {
            var address = new CellAddress(sheet.Id, row, column);
            if (_styles.TryGetValue(address, out var cached)) return cached;
            _evaluation.ResetBudget();
            var own = Raw(sheet, row, column); var style = own.Style;
            (int, string?) Edge(int nr, int nc, int side, int opposite, int kind, string? color)
            {
                if (nr < 0 || nc < 0 || nr >= WorkbookLimits.MaxRows || nc >= WorkbookLimits.MaxColumns) return (kind, color);
                var next = Raw(sheet, nr, nc); var other = next.Style;
                var otherKind = opposite switch { 0 => other.LeftBorder, 1 => other.RightBorder, 2 => other.TopBorder, _ => other.BottomBorder };
                var otherColor = opposite switch { 0 => other.LeftBorderColor, 1 => other.RightBorderColor, 2 => other.TopBorderColor, _ => other.BottomBorderColor };
                if (next.Ranks[opposite] < own.Ranks[side]) return (otherKind, otherColor);
                return (kind, color);
            }
            var left = Edge(row, column - 1, 0, 1, style.LeftBorder, style.LeftBorderColor);
            var right = Edge(row, column + 1, 1, 0, style.RightBorder, style.RightBorderColor);
            var top = Edge(row - 1, column, 2, 3, style.TopBorder, style.TopBorderColor);
            var bottom = Edge(row + 1, column, 3, 2, style.BottomBorder, style.BottomBorderColor);
            style = style with { LeftBorder = left.Item1, LeftBorderColor = left.Item2, RightBorder = right.Item1, RightBorderColor = right.Item2,
                TopBorder = top.Item1, TopBorderColor = top.Item2, BottomBorder = bottom.Item1, BottomBorderColor = bottom.Item2 };
            if (_styles.Count >= 16384) _styles.Clear();
            return _styles[address] = style;
        }
        private Projection Raw(WorksheetSnapshot sheet, int row, int column)
        {
            var address = new CellAddress(sheet.Id, row, column);
            if (_raw.TryGetValue(address, out var known)) return known;
            var baseline = TablePresentation.GetStyle(Snapshot, sheet, row, column);
            var ranks = new[] { baseline.LeftBorder, baseline.RightBorder, baseline.TopBorder, baseline.BottomBorder }.Select(v => v == 0 ? long.MaxValue : long.MaxValue - 1).ToArray();
            var matches = new List<(ConditionalFormatRule Rule, DifferentialStyle Style, ConditionalCellVisual? Visual)>();
            foreach (var rule in sheet.Data.ConditionalFormats.OrderBy(r => r.Priority).Where(r => r.Contains(row, column)))
            {
                bool match; var style = rule.Style; ConditionalCellVisual? visual = null;
                try
                {
                    if (rule.Kind >= ConditionalFormatKind.ColorScale) match = VisualMatch(sheet, new(row, column), rule, out style, out visual);
                    else match = Matches(sheet, new(row, column), rule);
                }
                catch (Exception ex) when (ex is NotSupportedException or FormatException or InvalidOperationException)
                { Errors.Add(sheet.Name + "：" + ex.Message); break; }
                if (!match) continue; matches.Add((rule, style, visual)); if (rule.StopIfTrue) break;
            }
            ConditionalBarVisual? bar = null; ConditionalIconVisual? icon = null; var hideBarValue = false; var hideIconValue = false;
            foreach (var item in matches.AsEnumerable().Reverse())
            {
                baseline = item.Style.Apply(baseline);
                var edges = new[] { item.Style.Left, item.Style.Right, item.Style.Top, item.Style.Bottom };
                for (var i = 0; i < edges.Length; i++) if (edges[i] is not null) ranks[i] = item.Rule.Priority;
                if (item.Visual?.Bar is { } b) { bar = b; hideBarValue = item.Visual.HideValue; }
                if (item.Visual?.Icon is { } c) { icon = c; hideIconValue = item.Visual.HideValue; }
            }
            if (_raw.Count >= 16384) _raw.Clear();
            return _raw[address] = new(baseline, ranks, bar is null && icon is null ? null : new(bar, icon, hideBarValue || hideIconValue));
        }
        private bool Matches(WorksheetSnapshot sheet, CellPosition position, ConditionalFormatRule rule)
        {
            static bool True(CellValue value) => value.Kind == CellValueKind.Boolean ? value.BooleanValue : value.Kind == CellValueKind.Number && value.NumberValue != 0;
            if (rule.Kind is not (ConditionalFormatKind.CellIs or ConditionalFormatKind.DuplicateValues or ConditionalFormatKind.UniqueValues))
                return True(_evaluation.Scalar(rule.Formula1!, sheet.Id, rule.Anchor, position));
            var value = _evaluation.Cell(new(sheet.Id, position.Row, position.Column));
            if (rule.Kind is ConditionalFormatKind.DuplicateValues or ConditionalFormatKind.UniqueValues)
            {
                if (value.Kind == CellValueKind.Blank) return false;
                var counts = Counts(sheet, rule); long count;
                if (value.Kind == CellValueKind.Text && value.TextValue.IndexOfAny(['*', '?']) >= 0)
                {
                    if (!counts.Wildcards.TryGetValue(value, out count))
                    {
                        try
                        {
                            var criterion = FormulaCriterion.Create(value, Snapshot.DateSystem);
                            count = counts.Texts.Where(p => criterion.Matches(p.Key)).Sum(p => p.Value);
                        }
                        catch (FormulaValueException) { return false; }
                        counts.Wildcards[value] = count;
                    }
                }
                else count = counts.Keys.GetValueOrDefault(Key(value));
                var duplicate = count > 1;
                return rule.Kind == ConditionalFormatKind.DuplicateValues ? duplicate : !duplicate;
            }
            var first = _evaluation.Scalar(rule.Formula1!, sheet.Id, rule.Anchor, position);
            bool Compare(string op, CellValue other) => True(_evaluation.Expression(new BinaryNode(op, new ValueNode(value), new ValueNode(other)), sheet.Id));
            return rule.Comparison switch
            {
                ConditionalComparison.Between => Compare(">=", first) && Compare("<=", _evaluation.Scalar(rule.Formula2!, sheet.Id, rule.Anchor, position)),
                ConditionalComparison.NotBetween => Compare("<", first) || Compare(">", _evaluation.Scalar(rule.Formula2!, sheet.Id, rule.Anchor, position)),
                ConditionalComparison.Equal => Compare("=", first), ConditionalComparison.NotEqual => Compare("<>", first),
                ConditionalComparison.LessThan => Compare("<", first), ConditionalComparison.LessThanOrEqual => Compare("<=", first),
                ConditionalComparison.GreaterThan => Compare(">", first), ConditionalComparison.GreaterThanOrEqual => Compare(">=", first), _ => false
            };
        }
        private string Key(CellValue value)
        {
            if (value.Kind == CellValueKind.Blank || value.Kind == CellValueKind.Text && value.TextValue.Length == 0) return "empty";
            if (value.Kind == CellValueKind.Number) return "n:" + value.NumberValue.ToString("R", CultureInfo.InvariantCulture);
            var text = CellInput.DisplayText(value).Normalize(NormalizationForm.FormKC);
            if (value.Kind == CellValueKind.Text && FormulaCriterion.TryNumber(text, Snapshot.DateSystem, out var number))
                return "n:" + number.ToString("R", CultureInfo.InvariantCulture);
            return value.Kind + ":" + text.ToUpperInvariant();
        }
        private Distribution Counts(WorksheetSnapshot sheet, ConditionalFormatRule rule)
        {
            if (_counts.TryGetValue((sheet.Id, rule.Priority), out var known)) return known;
            var allocation = rule.Ranges.Sum(range => (long)(range.Bottom - range.Top + 1) * (range.Right - range.Left + 1));
            if (allocation + _countedCells > 2_000_000)
                throw new NotSupportedException("本快照的重复值条件格式累计区域超过 2,000,000 格，请缩小规则范围。");
            _countedCells += allocation;
            var result = new Dictionary<string, long>();
            var texts = new Dictionary<CellValue, long>();
            var visited = new HashSet<CellPosition>();
            foreach (var range in rule.Ranges)
            {
                for (var row = range.Top; row <= range.Bottom; row++)
                for (var col = range.Left; col <= range.Right; col++)
                {
                    if (!visited.Add(new(row, col))) continue;
                    var value = _evaluation.Cell(new(sheet.Id, row, col)); var key = Key(value); result[key] = result.GetValueOrDefault(key) + 1;
                    if (value.Kind == CellValueKind.Text) texts[value] = texts.GetValueOrDefault(value) + 1;
                }
            }
            return _counts[(sheet.Id, rule.Priority)] = new(result, texts);
        }
    }
}
