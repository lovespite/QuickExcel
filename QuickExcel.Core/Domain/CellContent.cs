namespace QuickExcel.Domain;

public enum CellContentKind { Literal, Formula }

/// <summary>Immutable input content. Calculated results never replace the original formula.</summary>
public sealed record CellContent
{
    private readonly CellValue _literal;
    private readonly string? _formula;

    private CellContent(CellContentKind kind, CellValue literal, string? formula, int styleId)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(styleId);
        Kind = kind;
        _literal = literal;
        _formula = formula;
        StyleId = styleId;
    }

    public CellContentKind Kind { get; }
    /// <summary>Zero denotes the default style; the workbook will own the style table.</summary>
    public int StyleId { get; }
    public CellValue LiteralValue => Kind == CellContentKind.Literal ? _literal
        : throw new InvalidOperationException("Formula content has no literal value.");
    public string FormulaText => Kind == CellContentKind.Formula ? _formula!
        : throw new InvalidOperationException("Literal content has no formula.");

    public static CellContent FromLiteral(CellValue value, int styleId = 0) =>
        new(CellContentKind.Literal, value, null, styleId);

    public static CellContent FromFormula(string formulaText, int styleId = 0)
    {
        ArgumentNullException.ThrowIfNull(formulaText);
        if (formulaText.Length < 2 || formulaText[0] != '=' || formulaText.AsSpan(1).IsWhiteSpace())
            throw new ArgumentException("A formula must start with '=' and contain an expression.", nameof(formulaText));
        // Parsing and binding are calculation concerns. Preserve even unsupported/malformed expressions.
        return new(CellContentKind.Formula, default, formulaText, styleId);
    }

    public override string ToString() => Kind == CellContentKind.Literal
        ? $"Literal({_literal}, StyleId={StyleId})"
        : $"Formula({_formula}, StyleId={StyleId})";
}
