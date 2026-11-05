using QuickExcel.Core;
namespace QuickExcel.Domain;

/// <summary>Sheet operation restrictions, not encryption. True flags mean the operation is prohibited.</summary>
public sealed record WorksheetProtection
{
    public bool Sheet { get; init; } = true;
    public bool Objects { get; init; } = true;
    public bool Scenarios { get; init; } = true;
    public bool FormatCells { get; init; } = true;
    public bool FormatColumns { get; init; } = true;
    public bool FormatRows { get; init; } = true;
    public bool InsertColumns { get; init; } = true;
    public bool InsertRows { get; init; } = true;
    public bool InsertHyperlinks { get; init; } = true;
    public bool DeleteColumns { get; init; } = true;
    public bool DeleteRows { get; init; } = true;
    public bool SelectLockedCells { get; init; }
    public bool Sort { get; init; } = true;
    public bool AutoFilter { get; init; } = true;
    public bool PivotTables { get; init; } = true;
    public bool SelectUnlockedCells { get; init; }
    public string? Password { get; init; }
    public string? AlgorithmName { get; init; }
    public string? HashValue { get; init; }
    public string? SaltValue { get; init; }
    public uint? SpinCount { get; init; }
    public bool HasPassword => Password is not null || AlgorithmName is not null || HashValue is not null || SaltValue is not null || SpinCount is not null;
    internal void Validate()
    {
        if (Password is { } p && (p.Length != 4 || p.Any(c => !Uri.IsHexDigit(c))) ||
            AlgorithmName?.Length > 64 || HashValue?.Length > 1024 || SaltValue?.Length > 1024 || SpinCount > 10_000_000)
            throw new ArgumentException(MessageText.Get("工作表保护校验参数无效或超限。"));
    }
}
