namespace QuickExcel.Domain;

/// <summary>Capability of a specific feature, not a verdict on the entire workbook.</summary>
public enum CompatibilityKind { Editable, PreservedOnly, CannotCalculate, PotentialLoss, Rejected }

public sealed record CompatibilityDiagnostic(string Code, CompatibilityKind Kind, string Feature, string Message,
    string? Part = null, string? SheetName = null, string? Range = null,
    CompatibilityDiagnosticSummary? Summary = null, string? TypeNamespace = null)
{
    public bool IsReportNotice => Code.StartsWith("report.", StringComparison.Ordinal);
    public string TypeCode => Code == "xlsx.affected-location" ? "xlsx.unmodeled" : Code;
    public string TypeFeature
    {
        get
        {
            var separator = Feature.IndexOf(" · ", StringComparison.Ordinal);
            if (separator < 0) return Feature;
            var suffix = Feature[(separator + 3)..];
            return suffix.StartsWith("cellXfs[", StringComparison.Ordinal) ||
                suffix.StartsWith("cellStyleXfs[", StringComparison.Ordinal) ||
                suffix.StartsWith("sharedStrings[", StringComparison.Ordinal) ? Feature[..separator] : Feature;
        }
    }
}

public sealed record WorkbookCompatibilityReport(bool OriginalPackageUnchanged,
    IReadOnlyList<CompatibilityDiagnostic> Diagnostics);
