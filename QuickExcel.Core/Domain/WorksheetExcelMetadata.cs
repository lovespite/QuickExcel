using System.Xml;
using System.Xml.Linq;

namespace QuickExcel.Domain;

/// <summary>Validated inert font payload used by Excel's sheet-level phonetic display preference.</summary>
public sealed record WorksheetPhoneticProperties(string FontXml, string? Type = null, string? Alignment = null)
{
    private readonly string _fontXml = Canonicalize(FontXml);
    public string FontXml { get => _fontXml; init => _fontXml = Canonicalize(value); }

    // Namespace prefixes, attribute order and insignificant whitespace are not font identity.
    // Storage still validates the supported OOXML font shape before writing this inert payload.
    private static string Canonicalize(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml) || xml.Length > 16384)
            throw new ArgumentException("Invalid Excel phonetic font payload.", nameof(xml));
        try
        {
            using var input = new StringReader(xml);
            using var reader = XmlReader.Create(input, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16384,
                IgnoreWhitespace = true, IgnoreComments = true, IgnoreProcessingInstructions = true
            });
            var font = XElement.Load(reader);
            foreach (var element in font.DescendantsAndSelf())
            {
                var attributes = element.Attributes().Where(a => !a.IsNamespaceDeclaration)
                    .OrderBy(a => a.Name.NamespaceName, StringComparer.Ordinal).ThenBy(a => a.Name.LocalName, StringComparer.Ordinal)
                    .Select(a => new XAttribute(a)).ToArray();
                element.ReplaceAttributes(attributes);
            }
            var result = font.ToString(SaveOptions.DisableFormatting);
            if (result.Length > 16384) throw new ArgumentException("Excel phonetic font payload exceeds its budget.", nameof(xml));
            return result;
        }
        catch (XmlException exception) { throw new ArgumentException("Invalid Excel phonetic font XML.", nameof(xml), exception); }
    }

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(FontXml) || FontXml.Length > 16384 ||
            Type is not (null or "halfwidthKatakana" or "fullwidthKatakana" or "Hiragana" or "noConversion") ||
            Alignment is not (null or "noControl" or "left" or "center" or "distributed"))
            throw new ArgumentException("Invalid Excel phonetic-display properties.");
    }
}

/// <summary>Excel error-checking preferences attached to an explicit cell rectangle; no calculation semantics.</summary>
[Flags]
public enum IgnoredCellErrors
{
    None = 0,
    CalculatedColumn = 1 << 0,
    EvaluationError = 1 << 1,
    Formula = 1 << 2,
    FormulaRange = 1 << 3,
    ListDataValidation = 1 << 4,
    NumberStoredAsText = 1 << 5,
    TwoDigitTextYear = 1 << 6,
    UnlockedFormula = 1 << 7,
    EmptyCellReference = 1 << 8
}

public sealed record WorksheetIgnoredError(CellRectangle Range, IgnoredCellErrors Errors)
{
    public void Validate()
    {
        WorkbookLimits.ValidatePosition(Range.Top, Range.Left);
        WorkbookLimits.ValidatePosition(Range.Bottom, Range.Right);
        if (Range.Top > Range.Bottom || Range.Left > Range.Right || ((int)Errors & ~511) != 0)
            throw new ArgumentException("Invalid Excel ignored-error range or flags.");
    }
}
