using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Calculation;

/// <summary>Excel conditional-function criteria are distinct from ordinary formula comparisons.</summary>
internal sealed class FormulaCriterion
{
    private string _operator = "=";
    private CellValue _operand;
    private bool _emptyText;
    private Regex? _wildcard;
    private readonly ExcelDateSystem _dateSystem;
    private FormulaCriterion(ExcelDateSystem dateSystem) => _dateSystem = dateSystem;

    public static FormulaCriterion Create(CellValue value, ExcelDateSystem dateSystem)
    {
        var criterion = new FormulaCriterion(dateSystem);
        if (value.Kind != CellValueKind.Text)
        { criterion._operand = value.Kind == CellValueKind.Blank ? CellValue.FromNumber(0) : value; return criterion; }
        var text = value.TextValue;
        if (text.Length > 255) throw new FormulaValueException(CellErrorCode.Value, "条件文本超过 255 字符。");
        if (text.Length == 0) { criterion._operand = value; criterion._emptyText = true; return criterion; }
        var op = new[] { "<=", ">=", "<>", "=", "<", ">" }.FirstOrDefault(text.StartsWith);
        if (op is not null) { criterion._operator = op; text = text[op.Length..]; }
        if (text.Length == 0) { criterion._operand = CellValue.Blank; return criterion; }
        if (bool.TryParse(text, out var boolean)) criterion._operand = CellValue.FromBoolean(boolean);
        else if (TryError(text, out var error)) criterion._operand = CellValue.FromError(error);
        else if (TryNumber(text, dateSystem, out var number)) criterion._operand = CellValue.FromNumber(number);
        else
        {
            criterion._operand = CellValue.FromText(text);
            if (text.IndexOfAny(['*', '?']) >= 0 && criterion._operator is "=" or "<>")
            {
                var pattern = new StringBuilder("\\A");
                for (var i = 0; i < text.Length; i++)
                {
                    if (text[i] == '~' && i + 1 < text.Length && text[i + 1] is '*' or '?' or '~') pattern.Append(Regex.Escape(text[++i].ToString()));
                    else pattern.Append(text[i] switch { '*' => ".*", '?' => ".", _ => Regex.Escape(text[i].ToString()) });
                }
                pattern.Append("\\z");
                criterion._wildcard = new(pattern.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline | RegexOptions.NonBacktracking);
            }
        }
        return criterion;
    }

    public bool Matches(CellValue value)
    {
        if (_emptyText) return value.Kind == CellValueKind.Blank || value is { Kind: CellValueKind.Text, TextValue.Length: 0 };
        if (_operand.Kind == CellValueKind.Blank)
            return _operator switch { "=" => value.Kind == CellValueKind.Blank, "<>" => value.Kind != CellValueKind.Blank, _ => false };
        if (_operator == "=" && _operand.Kind == CellValueKind.Number && value.Kind == CellValueKind.Text && TryNumber(value.TextValue, _dateSystem, out var number))
            return number == _operand.NumberValue;
        if (value.Kind != _operand.Kind) return _operator == "<>";
        return value.Kind switch
        {
            CellValueKind.Number => Test(value.NumberValue.CompareTo(_operand.NumberValue)),
            CellValueKind.Boolean => Test(value.BooleanValue.CompareTo(_operand.BooleanValue)),
            CellValueKind.Error => Test(ErrorOrder(value.ErrorCode).CompareTo(ErrorOrder(_operand.ErrorCode))),
            CellValueKind.Text when _wildcard is not null => _operator == "<>" ^ _wildcard.IsMatch(value.TextValue),
            CellValueKind.Text => Test(string.Compare(value.TextValue, _operand.TextValue, StringComparison.InvariantCultureIgnoreCase)),
            _ => false
        };
    }

    private bool Test(int compare) => _operator switch { "=" => compare == 0, "<>" => compare != 0, "<" => compare < 0, "<=" => compare <= 0, ">" => compare > 0, _ => compare >= 0 };
    private static int ErrorOrder(CellErrorCode error) => error switch
    { CellErrorCode.Null => 0, CellErrorCode.DivisionByZero => 7, CellErrorCode.Value => 15, CellErrorCode.Reference => 23, CellErrorCode.Name => 29, CellErrorCode.Number => 36, CellErrorCode.NotAvailable => 42, _ => 99 };
    private static bool TryError(string text, out CellErrorCode error)
    {
        var parsed = text.ToUpperInvariant() switch { "#NULL!" => CellErrorCode.Null, "#DIV/0!" => CellErrorCode.DivisionByZero, "#VALUE!" => CellErrorCode.Value,
            "#REF!" => CellErrorCode.Reference, "#NAME?" => CellErrorCode.Name, "#NUM!" => CellErrorCode.Number, "#N/A" => CellErrorCode.NotAvailable, _ => (CellErrorCode?)null };
        error = parsed.GetValueOrDefault(); return parsed is not null;
    }
    private static bool TryNumber(string text, ExcelDateSystem system, out double number)
    {
        text = text.Trim(); var percent = text.EndsWith('%'); if (percent) text = text[..^1];
        var style = NumberStyles.Float;
        if (Regex.IsMatch(text, @"^[+-]?\d{1,3}(,\d{3})+(\.\d*)?([Ee][+-]?\d+)?$", RegexOptions.CultureInvariant)) style |= NumberStyles.AllowThousands;
        if (double.TryParse(text, style, CultureInfo.InvariantCulture, out number) && double.IsFinite(number)) { if (percent) number /= 100; return true; }
        // Stable year-first criteria dates; ambiguous locale-specific day/month strings remain text.
        if (!percent && DateTime.TryParseExact(text, ["yyyy/M/d", "yyyy-M-d", "yyyy/MM/dd", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            number = (date - (system == ExcelDateSystem.Mac1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31))).TotalDays;
            if (system == ExcelDateSystem.Windows1900 && date >= new DateTime(1900, 3, 1)) number++;
            return true;
        }
        return false;
    }
}
