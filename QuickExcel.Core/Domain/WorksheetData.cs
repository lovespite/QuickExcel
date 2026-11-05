using System.Collections.Immutable;
using QuickExcel.Core;

namespace QuickExcel.Domain;

public enum AutoFilterComparison { Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual }
public sealed record AutoFilterCondition(AutoFilterComparison Comparison, string Operand);
public sealed record AutoFilterColor(bool Font, string? Argb);
public sealed record AutoFilterTop(int Count, bool Percent = false, bool Bottom = false);

/// <summary>Excel checklist labels or one/two custom conditions. Column indexes are absolute model coordinates.</summary>
public sealed class AutoFilterColumn
{
    public AutoFilterColumn(IEnumerable<string>? values = null, bool includeBlank = false,
        AutoFilterCondition? first = null, AutoFilterCondition? second = null, bool useAnd = false, bool showButton = true, AutoFilterColor? color = null, AutoFilterTop? top = null)
    {
        if (top is not null && (color is not null || values is not null || includeBlank || first is not null || second is not null || top.Count < 1 || top.Count > (top.Percent ? 100 : 500)))
            throw new ArgumentException(MessageText.Get("前后项筛选需要独立的 1–500 项或 1–100 百分比。"));
        Top = top;
        if (color is not null && (values is not null || includeBlank || first is not null || second is not null ||
            color.Font && color.Argb is null || color.Argb is { } rgb && (rgb.Length != 8 || !rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase) || rgb.Any(c => !Uri.IsHexDigit(c)))))
            throw new ArgumentException(MessageText.Get("颜色筛选需要独立的不透明 ARGB 颜色，或无填充。"));
        Color = color is null ? null : color with { Argb = color.Argb?.ToUpperInvariant() };
        foreach (var condition in new[] { first, second }.OfType<AutoFilterCondition>())
            if (!Enum.IsDefined(condition.Comparison) || condition.Operand is null || condition.Operand.Length > 1024)
                throw new ArgumentException(MessageText.Get("筛选条件无效或超过 1024 个字符。"));
        // Excel canonicalizes one equality (or two OR equalities) to its display-value checklist.
        // A tilde only quotes * and ? in this conversion; a literal ~~ remains two characters.
        bool Label(AutoFilterCondition condition, out string label)
        {
            var text = new System.Text.StringBuilder(); label = "";
            if (condition.Comparison != AutoFilterComparison.Equal) return false;
            for (var i = 0; i < condition.Operand.Length; i++)
            {
                var c = condition.Operand[i];
                if (c == '~' && i + 1 < condition.Operand.Length && condition.Operand[i + 1] is '*' or '?') text.Append(condition.Operand[++i]);
                else if (c is '*' or '?') return false; else text.Append(c);
            }
            label = text.ToString();
            if (label.Length > 0 && string.IsNullOrWhiteSpace(label)) throw new ArgumentException(MessageText.Get("空格等值条件在 Excel 中不能稳定保存，请使用明确的值列表或空白选项。"));
            return true;
        }
        if (values is null && first is not null && Label(first, out var firstLabel) && (second is null || !useAnd && Label(second, out _)))
        {
            var labels = new List<string> { firstLabel };
            if (second is not null) { Label(second, out var secondLabel); labels.Add(secondLabel); }
            includeBlank = labels.Contains(""); values = labels.Where(v => v.Length > 0).ToArray(); first = second = null; useAnd = false;
        }
        Values = values?.Take(100_001).ToImmutableArray(); IncludeBlank = includeBlank;
        First = first; Second = second; UseAnd = useAnd; ShowButton = showButton;
        if (Values is { } list && (list.Length > 100_000 || list.Any(v => v is null || v.Length > WorkbookLimits.MaxCellTextLength)) ||
            Values is not null && First is not null || Second is not null && First is null || Values is null && includeBlank)
            throw new ArgumentException(MessageText.Get("筛选值列表或条件无效。"));
        if (Values is { IsEmpty: true } && !includeBlank) throw new ArgumentException(MessageText.Get("请至少选择一个值或空白；Excel 不支持可重应用的空值列表。"));
        if (Values?.Sum(s => (long)s.Length) > 8 * 1024 * 1024) throw new ArgumentException(MessageText.Get("筛选值文本超过 8 Mi 字符。"));
    }
    public ImmutableArray<string>? Values { get; }
    public bool IncludeBlank { get; }
    public AutoFilterCondition? First { get; }
    public AutoFilterCondition? Second { get; }
    public bool UseAnd { get; }
    public bool ShowButton { get; }
    public AutoFilterColor? Color { get; }
    public AutoFilterTop? Top { get; }
    public bool HasCriteria => Values is not null || First is not null || Color is not null || Top is not null;
    public AutoFilterColumn WithButton(bool showButton) => new(Values, IncludeBlank, First, Second, UseAnd, showButton, Color, Top);
    public bool ContentEquals(AutoFilterColumn other) => IncludeBlank == other.IncludeBlank && First == other.First && Second == other.Second &&
        Color == other.Color && Top == other.Top && UseAnd == other.UseAnd && ShowButton == other.ShowButton && (Values is null ? other.Values is null : other.Values is { } list && Values.Value.SequenceEqual(list));
}

public sealed class WorksheetAutoFilter
{
    public WorksheetAutoFilter(CellRectangle range, IEnumerable<KeyValuePair<int, AutoFilterColumn>>? columns = null)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Top > range.Bottom || range.Left > range.Right) throw new ArgumentException(MessageText.Get("筛选区域无效。"));
        Range = range; Columns = columns?.ToImmutableDictionary() ?? ImmutableDictionary<int, AutoFilterColumn>.Empty;
        if (Columns.Any(p => p.Key < range.Left || p.Key > range.Right || p.Value is null)) throw new ArgumentException(MessageText.Get("筛选列越界。"));
        if ((long)(range.Bottom - range.Top) * Columns.Count > 2_000_000) throw new ArgumentException(MessageText.Get("筛选条件遍历超过 2,000,000 格，请缩小区域或条件列数。"));
    }
    public CellRectangle Range { get; }
    public ImmutableDictionary<int, AutoFilterColumn> Columns { get; }
    public bool ContentEquals(WorksheetAutoFilter other) => Range == other.Range && Columns.Count == other.Columns.Count &&
        Columns.All(p => other.Columns.TryGetValue(p.Key, out var value) && p.Value.ContentEquals(value));
}

/// <summary>Last sort specification; rows already have the resulting order. Range excludes any header.</summary>
public sealed class WorksheetSortState
{
    public WorksheetSortState(CellRectangle range, IEnumerable<SortKey> keys, string sortMethod = "pinYin")
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        Range = range; Keys = keys.Take(33).ToImmutableArray(); SortMethod = sortMethod;
        if (range.Top > range.Bottom || range.Left > range.Right || Keys.Length is < 1 or > 32 ||
            Keys.Any(k => k.Column < range.Left || k.Column > range.Right) || sortMethod is not ("pinYin" or "stroke"))
            throw new ArgumentException(MessageText.Get("排序状态范围或键无效。"));
        foreach (var key in Keys) key.Validate();
        if (Keys.Select(k => k.CaseSensitive).Distinct().Count() > 1) throw new ArgumentException(MessageText.Get("Excel 排序状态要求各列使用相同的大小写规则。"));
        if (Keys.Any(k => k.CustomOrder.Any(v => v.Contains(',') || v.Contains('\r') || v.Contains('\n'))))
            throw new ArgumentException(MessageText.Get("可保存的自定义排序项不能包含逗号或换行。"));
    }
    public CellRectangle Range { get; }
    public ImmutableArray<SortKey> Keys { get; }
    public string SortMethod { get; }
    public bool ContentEquals(WorksheetSortState other) => Range == other.Range && SortMethod == other.SortMethod && Keys.Length == other.Keys.Length &&
        Keys.Zip(other.Keys).All(p => p.First.Column == p.Second.Column && p.First.Descending == p.Second.Descending &&
            p.First.CaseSensitive == p.Second.CaseSensitive && p.First.Basis == p.Second.Basis &&
            string.Equals(p.First.PreferredColor, p.Second.PreferredColor, StringComparison.OrdinalIgnoreCase) && p.First.CustomOrder.SequenceEqual(p.Second.CustomOrder));
}

public sealed record OutlineSettings(bool SummaryBelow = true, bool SummaryRight = true, bool ShowSymbols = true, bool ApplyStyles = false);

public sealed record WorksheetData(WorksheetAutoFilter? Filter = null, WorksheetSortState? Sort = null)
{
    public ImmutableArray<WorksheetComment> ThreadComments { get; init; } = [];
    public bool CommentFallback { get; init; }
    public WorksheetProtection? Protection { get; init; }
    public double? DefaultColumnWidth { get; init; }
    public bool ImplicitColumnWidth { get; init; }
    public ImmutableArray<WorksheetChart> Charts { get; init; } = [];
    public ImmutableArray<WorksheetPivot> Pivots { get; init; } = [];
    public OutlineSettings? Outline { get; init; }
    public ImmutableArray<WorksheetTable> Tables { get; init; } = [];
    public ImmutableArray<WorksheetValidation> Validations { get; init; } = [];
    public ImmutableArray<ConditionalFormatRule> ConditionalFormats { get; init; } = [];
    public IEnumerable<WorksheetAutoFilter> AllFilters => (Filter is null ? Enumerable.Empty<WorksheetAutoFilter>() : [Filter]).Concat(Tables.Select(t => t.Filter).OfType<WorksheetAutoFilter>());
    public bool IsEmpty => ThreadComments.IsEmpty && Filter is null && Sort is null && Tables.IsEmpty && Validations.IsEmpty && ConditionalFormats.IsEmpty && Outline is null && Protection is null && Charts.IsEmpty && Pivots.IsEmpty && DefaultColumnWidth is null;
    public bool ContentEquals(WorksheetData other, bool equivalentFormulas = false) => (Filter is null ? other.Filter is null : other.Filter is not null && Filter.ContentEquals(other.Filter)) &&
        ThreadComments.SequenceEqual(other.ThreadComments) && CommentFallback == other.CommentFallback && DefaultColumnWidth == other.DefaultColumnWidth && ImplicitColumnWidth == other.ImplicitColumnWidth && Charts.Length == other.Charts.Length && Charts.Zip(other.Charts).All(p => p.First.ContentEquals(p.Second)) && Protection == other.Protection && Outline == other.Outline && (Sort is null ? other.Sort is null : other.Sort is not null && Sort.ContentEquals(other.Sort)) &&
        Tables.Length == other.Tables.Length && Tables.Zip(other.Tables).All(p => p.First.ContentEquals(p.Second, equivalentFormulas)) &&
        Pivots.Length == other.Pivots.Length && Pivots.Zip(other.Pivots).All(p => p.First.ContentEquals(p.Second)) &&
        Validations.Length == other.Validations.Length && Validations.Zip(other.Validations).All(p => p.First.ContentEquals(p.Second)) &&
        ConditionalFormats.Length == other.ConditionalFormats.Length && ConditionalFormats.Zip(other.ConditionalFormats).All(p => p.First.ContentEquals(p.Second));
}
