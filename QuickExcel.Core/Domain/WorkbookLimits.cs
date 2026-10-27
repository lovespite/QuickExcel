namespace QuickExcel.Domain;

/// <summary>Product limits applied before snapshots enter a session.</summary>
public static class WorkbookLimits
{
    public const int MaxRows = 100_000;
    public const int MaxColumns = 1_024;
    public const int MaxCells = 1_000_000;
    public const int MaxSheets = 64;
    public const int MaxCellTextLength = 32_767;
    public const int MaxSheetNameLength = 31;

    public static void ValidatePosition(int row, int column)
    {
        if (row < 0 || row >= MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0 || column >= MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
    }

    public static void ValidateSheetName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (string.IsNullOrWhiteSpace(name) || name.Length > MaxSheetNameLength ||
            name.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0 ||
            name.Any(char.IsControl) || name[0] == '\'' || name[^1] == '\'')
            throw new ArgumentException("Sheet names must be nonblank, at most 31 characters, and valid Excel sheet names.", nameof(name));
    }

    public static void ValidateText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (text.Length > MaxCellTextLength)
            throw new ArgumentException($"Cell text cannot exceed {MaxCellTextLength} characters.", nameof(text));
    }

    internal static void ValidateDimension(double? value, string parameterName)
    {
        if (value.HasValue && (!double.IsFinite(value.Value) || value.Value < 0))
            throw new ArgumentOutOfRangeException(parameterName);
    }
}
