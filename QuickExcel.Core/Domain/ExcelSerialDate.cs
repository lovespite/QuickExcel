namespace QuickExcel.Domain;

public enum ExcelDateSystem { Windows1900, Mac1904 }

/// <summary>
/// An explicit interpretation of a numeric serial, not a sixth cell value type.
/// Negative serials extend the calendar before the epoch; time fractions have double precision.
/// </summary>
public readonly record struct ExcelSerialDate
{
    private static readonly DateTime Epoch1900 = new(1899, 12, 31);
    private static readonly DateTime Epoch1904 = new(1904, 1, 1);
    private static readonly DateTime March1900 = new(1900, 3, 1);

    public ExcelSerialDate(double serial, ExcelDateSystem dateSystem)
    {
        if (!double.IsFinite(serial)) throw new ArgumentOutOfRangeException(nameof(serial));
        ValidateSystem(dateSystem);
        Serial = serial;
        DateSystem = dateSystem;
    }

    public double Serial { get; }
    public ExcelDateSystem DateSystem { get; }
    public bool IsFictitiousLeapDay => DateSystem == ExcelDateSystem.Windows1900 && Serial >= 60 && Serial < 61;

    /// <summary>Returns false for the fictitious 1900-02-29 or dates outside the DateTime range.</summary>
    public bool TryToDateTime(out DateTime value)
    {
        value = default;
        if (IsFictitiousLeapDay) return false;
        var epoch = DateSystem == ExcelDateSystem.Windows1900 ? Epoch1900 : Epoch1904;
        var days = DateSystem == ExcelDateSystem.Windows1900 && Serial >= 61 ? Serial - 1 : Serial;
        try
        {
            value = epoch.AddDays(days);
            return true;
        }
        catch (ArgumentOutOfRangeException) { return false; }
    }

    /// <summary>Accepts a wall-clock date only; callers must explicitly resolve any time-zone semantics.</summary>
    public static ExcelSerialDate FromDateTime(DateTime value, ExcelDateSystem dateSystem)
    {
        ValidateSystem(dateSystem);
        if (value.Kind != DateTimeKind.Unspecified)
            throw new ArgumentException("Excel dates require DateTimeKind.Unspecified.", nameof(value));
        var epoch = dateSystem == ExcelDateSystem.Windows1900 ? Epoch1900 : Epoch1904;
        var serial = (value - epoch).TotalDays;
        if (dateSystem == ExcelDateSystem.Windows1900 && value >= March1900) serial += 1;
        var result = new ExcelSerialDate(serial, dateSystem);
        if (!result.TryToDateTime(out _))
            throw new ArgumentOutOfRangeException(nameof(value), "The date cannot be represented at double precision.");
        return result;
    }

    /// <summary>Preserves the calendar date when changing systems, and the exact serial within one system.</summary>
    public bool TryConvertTo(ExcelDateSystem dateSystem, out ExcelSerialDate value)
    {
        ValidateSystem(dateSystem);
        if (dateSystem == DateSystem)
        {
            value = this;
            return true;
        }
        value = default;
        if (!TryToDateTime(out _)) return false;
        // Before the fictitious leap day the epoch distance is 1461, not 1462.
        var serial = DateSystem == ExcelDateSystem.Windows1900
            ? Serial - (Serial < 60 ? 1461 : 1462)
            : Serial + (Serial < -1401 ? 1461 : 1462);
        var converted = new ExcelSerialDate(serial, dateSystem);
        if (!converted.TryToDateTime(out _)) return false;
        value = converted;
        return true;
    }

    private static void ValidateSystem(ExcelDateSystem dateSystem)
    {
        if (!Enum.IsDefined(dateSystem)) throw new ArgumentOutOfRangeException(nameof(dateSystem));
    }
}
