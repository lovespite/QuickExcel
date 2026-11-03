namespace QuickExcel.Domain;

public enum CellContentKind { Literal, Formula }

/// <summary>Immutable input content. Calculated results never replace the original formula.</summary>
public sealed record CellContent
{
    private readonly CellValue _literal;
    private readonly string? _formula;

    private CellContent(CellContentKind kind, CellValue literal, string? formula, int styleId, RichCellText? richText = null, bool dynamicArray = false)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(styleId);
        Kind = kind;
        _literal = literal;
        _formula = formula;
        StyleId = styleId;
        RichText = richText;
        DynamicArray = dynamicArray;
    }

    public CellContentKind Kind { get; }
    /// <summary>Zero denotes the default style; the workbook will own the style table.</summary>
    public int StyleId { get; }
    public RichCellText? RichText { get; }
    /// <summary>Dynamic array evaluation is distinct from legacy scalar/implicit-intersection formulas.</summary>
    public bool DynamicArray { get; }
    public CellValue LiteralValue => Kind == CellContentKind.Literal ? _literal
        : throw new InvalidOperationException("Formula content has no literal value.");
    public string FormulaText => Kind == CellContentKind.Formula ? _formula!
        : throw new InvalidOperationException("Literal content has no formula.");

    public static CellContent FromLiteral(CellValue value, int styleId = 0) =>
        new(CellContentKind.Literal, value, null, styleId);

    public static CellContent FromRichText(RichCellText text, int styleId = 0) =>
        new(CellContentKind.Literal, CellValue.FromText(text.Text), null, styleId, text);
    public CellContent WithStyle(int styleId) => new(Kind, _literal, _formula, styleId, RichText, DynamicArray);
    public CellContent WithFormula(string formula) => FromFormula(formula, StyleId, DynamicArray);

    public static CellContent FromFormula(string formulaText, int styleId = 0, bool dynamicArray = false)
    {
        ArgumentNullException.ThrowIfNull(formulaText);
        if (formulaText.Length < 2 || formulaText[0] != '=' || formulaText.AsSpan(1).IsWhiteSpace())
            throw new ArgumentException("A formula must start with '=' and contain an expression.", nameof(formulaText));
        // Parsing and binding are calculation concerns. Preserve even unsupported/malformed expressions.
        var decoded = QuickExcel.Formulas.ExcelFormulaNames.Decode(formulaText);
        try { dynamicArray |= QuickExcel.Formulas.DynamicArrayFunctions.UsesArrays(QuickExcel.Formulas.FormulaSyntax.Parse(decoded)); }
        catch (Exception error) when (error is FormatException or OverflowException) { }
        return new(CellContentKind.Formula, default, decoded, styleId, dynamicArray: dynamicArray);
    }

    public override string ToString() => Kind == CellContentKind.Literal
        ? $"Literal({_literal}, StyleId={StyleId})"
        : $"Formula({_formula}, StyleId={StyleId})";
}
