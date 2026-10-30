using QuickExcel.Domain;

namespace QuickExcel.Core;

public static partial class CellDisplayFormatter
{
    public static bool SupportsFilterLabel(CellValue value, StyleDefinition style)
    {
        if (value.Kind is CellValueKind.Blank or CellValueKind.Boolean or CellValueKind.Error) return true;
        var code = style.NumberFormatCode ?? BuiltIn(style.NumberFormatId);
        if (code is null) return style.NumberFormatId == 0;
        if (code.Equals("General", StringComparison.OrdinalIgnoreCase)) return true;
        // Layout padding depends on Excel's font metrics; never turn a fallback display into a different filter.
        if (code.Contains('_') || code.Contains('*')) return false;
        var format = Formats.GetValue(style, s => Parse(s.NumberFormatCode ?? BuiltIn(s.NumberFormatId)));
        if (format.Sections.Any(s => !s.Supported)) return false;
        if (value.Kind == CellValueKind.Text) return format.Sections.Length != 4 || format.Sections[3].Tokens.All(t => t.Kind is 'l' or '@');
        var section = format.Select(value.NumberValue);
        return section is not null && (section.HasDate || RenderNumber(Math.Abs(value.NumberValue), section, Invariant, out _) is not null);
    }
}
