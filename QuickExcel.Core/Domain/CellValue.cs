using System.Globalization;

namespace QuickExcel.Domain;

public enum CellValueKind { Blank, Text, Number, Boolean, Error }

/// <summary>Semantic error identity; diagnostics and display spelling belong to other layers.</summary>
public enum CellErrorCode { Null, DivisionByZero, Value, Reference, Name, Number, NotAvailable, Cycle, Script, Calculation }

/// <summary>An immutable scalar. The default value is Blank; no string parsing or coercion occurs here.</summary>
public readonly record struct CellValue
{
    private readonly string? _text;
    private readonly double _number;
    private readonly bool _boolean;
    private readonly CellErrorCode _error;

    private CellValue(CellValueKind kind, string? text = null, double number = 0,
        bool boolean = false, CellErrorCode error = default)
    {
        Kind = kind;
        _text = text;
        _number = number;
        _boolean = boolean;
        _error = error;
    }

    public CellValueKind Kind { get; }
    public static CellValue Blank => default;
    public string TextValue => Kind == CellValueKind.Text ? _text! : throw WrongKind(CellValueKind.Text);
    public double NumberValue => Kind == CellValueKind.Number ? _number : throw WrongKind(CellValueKind.Number);
    public bool BooleanValue => Kind == CellValueKind.Boolean ? _boolean : throw WrongKind(CellValueKind.Boolean);
    public CellErrorCode ErrorCode => Kind == CellValueKind.Error ? _error : throw WrongKind(CellValueKind.Error);

    public static CellValue FromText(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new(CellValueKind.Text, text: value);
    }

    public static CellValue FromNumber(double value) => double.IsFinite(value)
        ? new(CellValueKind.Number, number: value)
        : FromError(CellErrorCode.Number);

    public static CellValue FromBoolean(bool value) => new(CellValueKind.Boolean, boolean: value);

    public static CellValue FromError(CellErrorCode code)
    {
        if (!Enum.IsDefined(code)) throw new ArgumentOutOfRangeException(nameof(code));
        return new(CellValueKind.Error, error: code);
    }

    /// <summary>Diagnostic notation only; UI formatting and edit serialization are separate contracts.</summary>
    public override string ToString() => Kind switch
    {
        CellValueKind.Blank => "Blank",
        CellValueKind.Text => $"Text({_text})",
        CellValueKind.Number => $"Number({_number.ToString("R", CultureInfo.InvariantCulture)})",
        CellValueKind.Boolean => $"Boolean({_boolean})",
        CellValueKind.Error => $"Error({_error})",
        _ => throw new InvalidOperationException("Unknown cell value kind.")
    };

    private InvalidOperationException WrongKind(CellValueKind expected) =>
        new($"Cannot read {Kind} as {expected}.");
}
