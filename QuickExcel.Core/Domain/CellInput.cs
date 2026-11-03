using System.Globalization;

namespace QuickExcel.Domain;

/// <summary>Explicit user-input parsing, separate from display and storage value interpretation.</summary>
public static class CellInput
{
    /// <summary>ISO date entry is enabled only by an explicit date-validation rule. Typed clipboard/storage values stay typed.</summary>
    public static CellContent ParseForCell(string text, WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column)
    {
        var style = sheet.GetStyleId(row, column);
        if (sheet.Data.Validations.Any(v => v.Kind == ValidationKind.Date && v.Contains(new(row, column))))
        {
            if (text == "1900-02-29" && snapshot.DateSystem == ExcelDateSystem.Windows1900)
                return CellContent.FromLiteral(CellValue.FromNumber(60), style);
            if (DateTime.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                return CellContent.FromLiteral(CellValue.FromNumber(ExcelSerialDate.FromDateTime(date, snapshot.DateSystem).Serial), style);
        }
        return Parse(text, style);
    }

    public static CellContent Parse(string text, int styleId = 0)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentOutOfRangeException.ThrowIfNegative(styleId);
        if (text.Length == 0) return CellContent.FromLiteral(CellValue.Blank, styleId);
        if (text[0] == '\'')
        {
            WorkbookLimits.ValidateText(text[1..]);
            return CellContent.FromLiteral(CellValue.FromText(text[1..]), styleId);
        }
        WorkbookLimits.ValidateText(text);
        if (text[0] == '=' && text.Length > 1 && !text.AsSpan(1).IsWhiteSpace())
            return CellContent.FromFormula(text, styleId);
        var trimmed = text.AsSpan().Trim();
        if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
            return CellContent.FromLiteral(CellValue.FromBoolean(true), styleId);
        if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            return CellContent.FromLiteral(CellValue.FromBoolean(false), styleId);
        if (!HasLeadingZeroes(trimmed) &&
            double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
            return CellContent.FromLiteral(CellValue.FromNumber(number), styleId);
        return CellContent.FromLiteral(CellValue.FromText(text), styleId);
    }

    /// <summary>Serializes input for an editor; typed text is escaped when parsing could reinterpret it.</summary>
    public static string EditText(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column)
    {
        if (QuickExcel.Core.ProtectionSafety.HidesFormula(snapshot, sheet, row, column)) return "";
        var cell = sheet.GetCell(row,column);
        if (cell?.SpillOwner is { } owner)
        {
            if (owner.SheetId != sheet.Id || QuickExcel.Core.ProtectionSafety.HidesFormula(snapshot,sheet,owner.Row,owner.Column)) return "";
            cell = sheet.GetCell(owner.Row,owner.Column);
        }
        return cell is null ? "" : EditText(cell.Content);
    }

    public static string EditText(CellContent content)
    {
        ArgumentNullException.ThrowIfNull(content);
        if (content.Kind == CellContentKind.Formula) return content.FormulaText;
        var value = content.LiteralValue;
        if (value.Kind != CellValueKind.Text) return DisplayText(value);
        var text = value.TextValue;
        var parsed = Parse(text);
        return parsed.Kind == CellContentKind.Literal && parsed.LiteralValue == value ? text : "'" + text;
    }

    public static string DisplayText(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => "",
        CellValueKind.Text => value.TextValue,
        CellValueKind.Number => value.NumberValue.ToString("R", CultureInfo.InvariantCulture),
        CellValueKind.Boolean => value.BooleanValue ? "TRUE" : "FALSE",
        CellValueKind.Error => value.ErrorCode switch
        {
            CellErrorCode.Null => "#NULL!",
            CellErrorCode.DivisionByZero => "#DIV/0!",
            CellErrorCode.Value => "#VALUE!",
            CellErrorCode.Reference => "#REF!",
            CellErrorCode.Name => "#NAME?",
            CellErrorCode.Number => "#NUM!",
            CellErrorCode.NotAvailable => "#N/A",
            CellErrorCode.Cycle => "#CYCLE!",
            CellErrorCode.Script => "#SCRIPT!",
            CellErrorCode.Calculation => "#CALC!",
            CellErrorCode.Spill => "#SPILL!",
            _ => throw new InvalidOperationException("Unknown cell error code.")
        },
        _ => throw new InvalidOperationException("Unknown cell value kind.")
    };

    /// <summary>Interprets an explicitly error-typed storage value. Unknown strings remain text.</summary>
    public static CellValue FromErrorText(string text) => TryFromErrorText(text, out var value)
        ? value : CellValue.FromText(text);

    public static bool TryFromErrorText(string text, out CellValue value)
    {
        ArgumentNullException.ThrowIfNull(text);
        CellErrorCode? code = text.ToUpperInvariant() switch
        {
            "#NULL!" => CellErrorCode.Null,
            "#DIV/0!" => CellErrorCode.DivisionByZero,
            "#VALUE!" => CellErrorCode.Value,
            "#REF!" => CellErrorCode.Reference,
            "#NAME?" => CellErrorCode.Name,
            "#NUM!" => CellErrorCode.Number,
            "#N/A" => CellErrorCode.NotAvailable,
            "#CYCLE!" => CellErrorCode.Cycle,
            "#SCRIPT!" => CellErrorCode.Script,
            "#CALC!" => CellErrorCode.Calculation,
            "#SPILL!" => CellErrorCode.Spill,
            _ => null
        };
        value = code.HasValue ? CellValue.FromError(code.Value) : CellValue.Blank;
        return code.HasValue;
    }

    private static bool HasLeadingZeroes(ReadOnlySpan<char> text)
    {
        if (!text.IsEmpty && text[0] is '+' or '-') text = text[1..];
        return text.Length > 1 && text[0] == '0' && char.IsAsciiDigit(text[1]);
    }
}
