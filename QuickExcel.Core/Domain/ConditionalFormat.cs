using QuickExcel.Core;
using System.Collections.Immutable;
using QuickExcel.Formulas;

namespace QuickExcel.Domain;

public enum ConditionalFormatKind { CellIs, Expression, DuplicateValues, UniqueValues, ContainsText, NotContainsText, BeginsWith, EndsWith, ColorScale, DataBar, IconSet, TimePeriod }
public enum ConditionalComparison { Between, NotBetween, Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual }

/// <summary>Nullable properties mean the rule does not change that property; explicit false remains meaningful.</summary>
public sealed record DifferentialStyle(bool? Bold = null, bool? Italic = null, bool? Strike = null, string? Underline = null,
    string? FontColor = null, string? FillColor = null, TableBorder? Left = null, TableBorder? Right = null, TableBorder? Top = null, TableBorder? Bottom = null,
    int? NumberFormatId = null, string? NumberFormatCode = null)
{
    internal void Validate()
    {
        new TableStyleElement("wholeTable", Bold: Bold, FontColor: FontColor, FillColor: FillColor, Left: Left, Right: Right, Top: Top, Bottom: Bottom,
            NumberFormatId: NumberFormatId, NumberFormatCode: NumberFormatCode).Validate();
        if (Underline is not (null or "none" or "single" or "double" or "singleAccounting" or "doubleAccounting"))
            throw new ArgumentException(MessageText.Get("条件格式下划线无效。"));
    }
    public StyleDefinition Apply(StyleDefinition original) => original with
    {
        Bold = Bold ?? original.Bold, Italic = Italic ?? original.Italic, Strike = Strike ?? original.Strike, Underline = Underline ?? original.Underline,
        FontColor = FontColor ?? original.FontColor, FillColor = FillColor ?? original.FillColor,
        LeftBorder = Left?.Style ?? original.LeftBorder, LeftBorderColor = Left is null ? original.LeftBorderColor : Left.Color,
        RightBorder = Right?.Style ?? original.RightBorder, RightBorderColor = Right is null ? original.RightBorderColor : Right.Color,
        TopBorder = Top?.Style ?? original.TopBorder, TopBorderColor = Top is null ? original.TopBorderColor : Top.Color,
        BottomBorder = Bottom?.Style ?? original.BottomBorder, BottomBorderColor = Bottom is null ? original.BottomBorderColor : Bottom.Color,
        NumberFormatId = NumberFormatId ?? original.NumberFormatId, NumberFormatCode = NumberFormatId.HasValue ? NumberFormatCode : original.NumberFormatCode
    };
}

public sealed class ConditionalFormatRule
{
    public Guid ScriptId { get; init; } = Guid.NewGuid();
    public ConditionalFormatRule(IEnumerable<CellRectangle> ranges, ConditionalFormatKind kind, int priority, DifferentialStyle style,
        string? formula1 = null, string? formula2 = null, ConditionalComparison comparison = ConditionalComparison.Between,
        bool stopIfTrue = false, string? text = null, ConditionalColorScale? colorScale = null, ConditionalDataBar? dataBar = null, ConditionalIconSet? iconSet = null, ConditionalDatePeriod? datePeriod = null)
    {
        if ((kind == ConditionalFormatKind.TimePeriod) != datePeriod.HasValue || datePeriod is { } period && !Enum.IsDefined(period) ||
            kind == ConditionalFormatKind.TimePeriod && (formula1 is not null || formula2 is not null || text is not null))
            throw new ArgumentException(MessageText.Get("发生日期规则必须选择支持的日期范围，不接受自定义公式或文本条件。"));
        DatePeriod = datePeriod;
        Ranges = ranges.Take(1025).ToImmutableArray();
        if (Ranges.Length is < 1 or > 1024 || priority < 1 || !Enum.IsDefined(kind) || !Enum.IsDefined(comparison) || text?.Length > 255)
            throw new ArgumentException(MessageText.Get("条件格式区域、优先级或规则类型无效。"));
        foreach (var range in Ranges)
        {
            WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
            if (range.Top > range.Bottom || range.Left > range.Right) throw new ArgumentException(MessageText.Get("条件格式区域无效。"));
        }
        Style = style ?? throw new ArgumentNullException(nameof(style)); Style.Validate();
        if (kind is ConditionalFormatKind.CellIs or ConditionalFormatKind.Expression && formula1 is null ||
            kind == ConditionalFormatKind.CellIs && comparison is ConditionalComparison.Between or ConditionalComparison.NotBetween && formula2 is null ||
            kind is ConditionalFormatKind.ContainsText or ConditionalFormatKind.NotContainsText or ConditionalFormatKind.BeginsWith or ConditionalFormatKind.EndsWith && (text is null || formula1 is null))
            throw new ArgumentException(MessageText.Get("条件格式公式或文本条件缺失。"));
        if ((kind == ConditionalFormatKind.ColorScale) != (colorScale is not null) || (kind == ConditionalFormatKind.DataBar) != (dataBar is not null) ||
            (kind == ConditionalFormatKind.IconSet) != (iconSet is not null)) throw new ArgumentException(MessageText.Get("条件格式类型和视觉规则不一致。"));
        dataBar?.Validate(); ColorScale = colorScale; DataBar = dataBar; IconSet = iconSet;
        if (kind is >= ConditionalFormatKind.ColorScale and <= ConditionalFormatKind.IconSet && (formula1 is not null || formula2 is not null || text is not null)) throw new ArgumentException(MessageText.Get("视觉规则使用阈值公式。"));
        if (kind is >= ConditionalFormatKind.ColorScale and <= ConditionalFormatKind.IconSet && style != new DifferentialStyle()) throw new NotSupportedException(MessageText.Get("视觉规则不能同时包含差异样式。"));
        if (kind is >= ConditionalFormatKind.ColorScale and <= ConditionalFormatKind.IconSet && stopIfTrue) throw new NotSupportedException(MessageText.Get("数据条、色阶和图标集不支持为真时停止。"));
        foreach (var formula in new[] { formula1, formula2 }.OfType<string>()) ValidateFormula(formula);
        Kind = kind; Priority = priority; Formula1 = Normalize(formula1); Formula2 = Normalize(formula2); Comparison = comparison; StopIfTrue = stopIfTrue; Text = text;
    }
    private static string? Normalize(string? formula) => formula is null ? null : "=" + FormulaSyntax.Print(FormulaSyntax.Parse(formula));
    internal static void ValidateFormula(string formula)
    {
        if (formula.Length > 8192) throw new ArgumentException(MessageText.Get("条件格式公式超过 8,192 字符。"));
        var expression = FormulaSyntax.Parse(formula);
        if (FormulaSyntax.Walk(expression).OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) || f.Name is "NOW" or "TODAY" or "RAND" or "RANDBETWEEN"))
            throw new NotSupportedException(MessageText.Get("条件格式仅支持已建模的非易变标准函数。"));
    }
    public ConditionalDatePeriod? DatePeriod { get; }
    public bool IsVisual => Kind is >= ConditionalFormatKind.ColorScale and <= ConditionalFormatKind.IconSet;
    public ConditionalColorScale? ColorScale { get; }
    public ConditionalDataBar? DataBar { get; }
    public ConditionalIconSet? IconSet { get; }
    public IEnumerable<ConditionalThreshold> Thresholds => ColorScale?.Thresholds.AsEnumerable() ?? IconSet?.Thresholds.AsEnumerable() ??
        (DataBar is { } bar ? new[] { bar.Minimum, bar.Maximum } : []);
    public ImmutableArray<CellRectangle> Ranges { get; }
    public ConditionalFormatKind Kind { get; }
    public int Priority { get; }
    public DifferentialStyle Style { get; }
    public string? Formula1 { get; }
    public string? Formula2 { get; }
    public ConditionalComparison Comparison { get; }
    public bool StopIfTrue { get; }
    public string? Text { get; }
    public CellPosition Anchor => new(Ranges[0].Top, Ranges[0].Left);
    public bool Contains(int row, int column) => Ranges.Any(r => r.Contains(row, column));
    public ConditionalFormatRule WithPriority(int priority) => new(Ranges, Kind, priority, Style, Formula1, Formula2, Comparison, StopIfTrue, Text, ColorScale, DataBar, IconSet, DatePeriod) { ScriptId = ScriptId };
    public ConditionalFormatRule WithRanges(IEnumerable<CellRectangle> ranges) => new(ranges, Kind, Priority, Style, Formula1, Formula2, Comparison, StopIfTrue, Text, ColorScale, DataBar, IconSet, DatePeriod) { ScriptId = ScriptId };
    public ConditionalFormatRule MapFormulas(Func<string, string> transform) => new(Ranges, Kind, Priority, Style,
        Formula1 is { } a ? transform(a) : null, Formula2 is { } b ? transform(b) : null, Comparison, StopIfTrue, Text,
        ColorScale?.Map(transform), DataBar?.Map(transform), IconSet?.Map(transform), DatePeriod) { ScriptId = ScriptId };
    public void ValidateBinding(WorkbookSnapshot snapshot, SheetId sheet)
    {
        foreach (var text in new[] { Formula1, Formula2 }.Concat(Thresholds.Select(t => t.Formula)).OfType<string>())
        {
            var expression = FormulaBinding.Bind(FormulaSyntax.Parse(text), snapshot, sheet, Anchor);
            if (FormulaSyntax.Walk(expression).OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) || f.Name is "NOW" or "TODAY" or "RAND" or "RANDBETWEEN"))
                throw new NotSupportedException(MessageText.Get("条件格式名称依赖了未支持或易变函数。"));
        }
    }
    public bool ContentEquals(ConditionalFormatRule other) => Ranges.SequenceEqual(other.Ranges) && Kind == other.Kind && Priority == other.Priority &&
        DatePeriod == other.DatePeriod && Style == other.Style && Formula1 == other.Formula1 && Formula2 == other.Formula2 && Comparison == other.Comparison && StopIfTrue == other.StopIfTrue && Text == other.Text &&
        (DataBar is null ? other.DataBar is null : other.DataBar is not null && DataBar.ContentEquals(other.DataBar)) && (ColorScale is null ? other.ColorScale is null : other.ColorScale is not null && ColorScale.ContentEquals(other.ColorScale)) &&
        (IconSet is null ? other.IconSet is null : other.IconSet is not null && IconSet.ContentEquals(other.IconSet));
    public long EstimatedBytes => 512L + Ranges.Length * 32L + new[] { Formula1, Formula2, Text, Style.NumberFormatCode }.Sum(s => (s?.Length ?? 0) * 2L) +
        Thresholds.Sum(t => 128L + (t.Formula?.Length ?? 0) * 2L) + (IsVisual ? 512 : 0);
}
