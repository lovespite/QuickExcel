namespace QuickExcel.Domain;

public enum PreservedFormulaKind { Array, DynamicArray, StructuredReference, ExternalReference, Shared, UnsupportedSyntax, DefinedName }

/// <summary>Read-only projection of an imported formula whose dependency/output semantics are not modeled.
/// Exact formula XML and cached result cells remain in WorkbookFileState.OriginalWorkbook.</summary>
public sealed record PreservedFormula(PreservedFormulaKind Kind, string Part, string SheetName,
    string Cell, string? Range, string Source, bool HasCompleteCache = true);
