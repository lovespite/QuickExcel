using QuickExcel.Core;
using System.Collections.Immutable;
using QuickExcel.Formulas;

namespace QuickExcel.Domain;

public enum ValidationKind { None, Whole, Decimal, List, Date, TextLength, Custom }
public enum ValidationOperator { Between, NotBetween, Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual }
public enum ValidationAlert { Stop, Warning, Information }

/// <summary>Formula coordinates are relative to the first sqref area's top-left cell.</summary>
public sealed class WorksheetValidation
{
    public WorksheetValidation(IEnumerable<CellRectangle> ranges, ValidationKind kind, string? formula1 = null, string? formula2 = null,
        ValidationOperator comparison = ValidationOperator.Between, bool allowBlank = true, bool showDropDown = true,
        bool showInput = true, bool showError = true, ValidationAlert alert = ValidationAlert.Stop,
        string? inputTitle = null, string? inputMessage = null, string? errorTitle = null, string? errorMessage = null)
    {
        Ranges = ranges.Take(1025).ToImmutableArray();
        if (Ranges.Length is < 1 or > 1024 || !Enum.IsDefined(kind) || !Enum.IsDefined(comparison) || !Enum.IsDefined(alert))
            throw new ArgumentException(MessageText.Get("数据验证类型或区域数无效。"));
        foreach (var r in Ranges)
        {
            WorkbookLimits.ValidatePosition(r.Top, r.Left); WorkbookLimits.ValidatePosition(r.Bottom, r.Right);
            if (r.Top > r.Bottom || r.Left > r.Right) throw new ArgumentException(MessageText.Get("数据验证区域无效。"));
        }
        if (inputTitle?.Length > 32 || errorTitle?.Length > 32 || inputMessage?.Length > 255 || errorMessage?.Length > 255)
            throw new ArgumentException(MessageText.Get("验证提示标题最多 32 字符，正文最多 255 字符。"));
        foreach (var formula in new[] { formula1, formula2 }.OfType<string>())
        {
            if (formula.Length is < 1 or > 8192) throw new ArgumentException(MessageText.Get("验证公式长度无效。"));
            var syntax = FormulaSyntax.Parse(formula);
            if (FormulaSyntax.Walk(syntax).OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) || f.Name is "NOW" or "TODAY" or "RAND" or "RANDBETWEEN"))
                throw new NotSupportedException(MessageText.Get("验证公式仅支持已建模的非易变标准函数。"));
        }
        if (kind != ValidationKind.None && formula1 is null ||
            kind is ValidationKind.Whole or ValidationKind.Decimal or ValidationKind.Date or ValidationKind.TextLength &&
                comparison is ValidationOperator.Between or ValidationOperator.NotBetween && formula2 is null)
            throw new ArgumentException(MessageText.Get("验证条件公式缺失。"));
        if (kind == ValidationKind.List && formula1 is not null && FormulaSyntax.Parse(formula1) is ValueNode { Value.Kind: CellValueKind.Text } literal &&
            (literal.Value.TextValue.Length > 255 || literal.Value.TextValue.Contains('\n') || literal.Value.TextValue.Contains('\r')))
            throw new NotSupportedException(MessageText.Get("直接列表最长 255 字符，不能含换行；较长列表请使用单行或单列范围。"));
        Kind = kind; Formula1 = formula1 is null ? null : "=" + FormulaSyntax.Print(FormulaSyntax.Parse(formula1));
        Formula2 = formula2 is null ? null : "=" + FormulaSyntax.Print(FormulaSyntax.Parse(formula2));
        Comparison = comparison; AllowBlank = allowBlank; ShowDropDown = showDropDown;
        ShowInput = showInput; ShowError = showError; Alert = alert; InputTitle = inputTitle; InputMessage = inputMessage; ErrorTitle = errorTitle; ErrorMessage = errorMessage;
    }
    public ImmutableArray<CellRectangle> Ranges { get; }
    public ValidationKind Kind { get; }
    public string? Formula1 { get; }
    public string? Formula2 { get; }
    public ValidationOperator Comparison { get; }
    public bool AllowBlank { get; }
    public bool ShowDropDown { get; }
    public bool ShowInput { get; }
    public bool ShowError { get; }
    public ValidationAlert Alert { get; }
    public string? InputTitle { get; }
    public string? InputMessage { get; }
    public string? ErrorTitle { get; }
    public string? ErrorMessage { get; }
    public CellPosition Anchor => new(Ranges[0].Top, Ranges[0].Left);
    public bool Contains(CellPosition position) => Ranges.Any(r => r.Contains(position.Row, position.Column));
    public void ValidateBinding(WorkbookSnapshot snapshot, SheetId sheet)
    {
        foreach (var text in new[] { Formula1, Formula2 }.OfType<string>())
        {
            var syntax = FormulaSyntax.Parse(text); var list = Kind == ValidationKind.List && text == Formula1;
            var bound = FormulaBinding.Bind(list ? new FunctionNode("SUM", [syntax]) : syntax, snapshot, sheet, Anchor);
            if (list) bound = ((FunctionNode)bound).Arguments[0];
            if (FormulaSyntax.Walk(bound).OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) || f.Name is "NOW" or "TODAY" or "RAND" or "RANDBETWEEN"))
                throw new NotSupportedException(MessageText.Get("验证名称间接依赖了未支持或易变函数。"));
            if (list && bound is not ValueNode { Value.Kind: CellValueKind.Text or CellValueKind.Error } &&
                (bound is not ReferenceNode r || r.Row != r.EndRow && r.Column != r.EndColumn ||
                    (long)(r.EndRow - r.Row + 1) * (r.EndColumn - r.Column + 1) > 10000))
                throw new NotSupportedException(MessageText.Get("列表来源须为不超过 10,000 格的单行或单列范围。"));
        }
    }
    public WorksheetValidation MapFormulas(Func<string, string> transform) => new(Ranges, Kind,
        Formula1 is { } a ? transform(a) : null, Formula2 is { } b ? transform(b) : null, Comparison, AllowBlank, ShowDropDown,
        ShowInput, ShowError, Alert, InputTitle, InputMessage, ErrorTitle, ErrorMessage);
    public bool ContentEquals(WorksheetValidation other) => Ranges.SequenceEqual(other.Ranges) && Kind == other.Kind && Formula1 == other.Formula1 &&
        Formula2 == other.Formula2 && Comparison == other.Comparison && AllowBlank == other.AllowBlank && ShowDropDown == other.ShowDropDown &&
        ShowInput == other.ShowInput && ShowError == other.ShowError && Alert == other.Alert && InputTitle == other.InputTitle &&
        InputMessage == other.InputMessage && ErrorTitle == other.ErrorTitle && ErrorMessage == other.ErrorMessage;
    public long EstimatedBytes => 256L + Ranges.Length * 32L + new[] { Formula1, Formula2, InputTitle, InputMessage, ErrorTitle, ErrorMessage }.Sum(t => (t?.Length ?? 0) * 2L);
}
