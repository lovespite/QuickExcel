namespace QuickExcel.Core;

public static partial class CellDisplayFormatter
{
    /// <summary>Normalize localized UI codes to SpreadsheetML and reject formats that would silently fall back.</summary>
    public static bool TryNormalizeNumberFormat(string code, out string normalized)
    {
        normalized = code;
        if (code.Length > 1024) return false;
        normalized = LocalTextCode(code, out var invalid, numberFormat: true);
        if (invalid) return false;
        var parsed = Parse(normalized);
        if (parsed.Sections.Length == 0 || parsed.Sections.Any(s => !s.Supported)) return false;
        for (var i = 0; i < parsed.Sections.Length; i++)
        {
            var section = parsed.Sections[i];
            if (i == 3)
            { if (section.Tokens.Any(t => t.Kind is not ('l' or '@' or '*'))) return false; }
            else if (!section.HasDate && RenderNumber(1.25, section, Invariant, out _) is null) return false;
        }
        return true;
    }
}
