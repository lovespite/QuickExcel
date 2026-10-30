namespace QuickExcel.Domain;

/// <summary>Capability of a specific feature, not a verdict on the entire workbook.</summary>
public enum CompatibilityKind { Editable, PreservedOnly, CannotCalculate, PotentialLoss, Rejected }

public sealed record CompatibilityDiagnostic(string Code, CompatibilityKind Kind, string Feature, string Message,
    string? Part = null, string? SheetName = null, string? Range = null);

public sealed record WorkbookCompatibilityReport(bool OriginalPackageUnchanged,
    IReadOnlyList<CompatibilityDiagnostic> Diagnostics);
