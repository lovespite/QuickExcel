using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

/// <summary>Explicit ISO Strict namespace/relationship mapping. Text values and external targets are never rewritten.</summary>
internal static class XlsxStrict
{
    internal const string StrictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string StrictRel = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private const string StrictDrawing = "http://purl.oclc.org/ooxml/drawingml/main";
    private const string Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly HashSet<string> PercentageValues = "tint shade alpha alphaOff alphaMod hueMod sat satOff satMod lum lumOff lumMod red redOff redMod green greenOff greenMod blue blueOff blueMod".Split(' ').ToHashSet(StringComparer.Ordinal);
    private static readonly IReadOnlyDictionary<string, string> Namespaces = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        [StrictMain] = Main.NamespaceName,
        [StrictRel] = Rel.NamespaceName,
        ["http://purl.oclc.org/ooxml/drawingml/main"] = "http://schemas.openxmlformats.org/drawingml/2006/main",
        ["http://purl.oclc.org/ooxml/drawingml/chart"] = "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ["http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"] = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        ["http://purl.oclc.org/ooxml/drawingml/chartDrawing"] = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        ["http://purl.oclc.org/ooxml/officeDocument/extendedProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        ["http://purl.oclc.org/ooxml/officeDocument/customProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        ["http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"] = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        ["http://purl.oclc.org/ooxml/officeDocument/sharedTypes"] = "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes"
    };
    internal static bool Map(XDocument document, bool toStrict, Action check)
    {
        var changed = false;
        var map = toStrict ? Namespaces.ToDictionary(p => p.Value, p => p.Key, StringComparer.Ordinal) : Namespaces;
        foreach (var element in document.Descendants())
        {
            check();
            if (element.Name.NamespaceName == StrictDrawing || toStrict && element.Name.NamespaceName == Drawing)
            foreach (var attribute in element.Attributes().Where(a => a.Name.NamespaceName.Length == 0 &&
                (a.Name.LocalName == "val" && PercentageValues.Contains(element.Name.LocalName) || a.Name.LocalName == "pos" && element.Name.LocalName == "gs" ||
                 a.Name.LocalName == "lim" && element.Name.LocalName == "miter" ||
                 a.Name.LocalName is "l" or "t" or "r" or "b" && element.Name.LocalName is "srcRect" or "fillRect" or "tileRect")))
            {
                if (!toStrict && attribute.Value.EndsWith('%'))
                {
                    if (!decimal.TryParse(attribute.Value[..^1], NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var percent) ||
                        percent < int.MinValue / 1000m || percent > int.MaxValue / 1000m || percent * 1000 != decimal.Truncate(percent * 1000))
                        throw new NotSupportedException("Strict DrawingML 百分比超出支持精度或范围。");
                    attribute.Value = (percent * 1000).ToString("0", CultureInfo.InvariantCulture); changed = true;
                }
                else if (toStrict && int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var units))
                { attribute.Value = (units / 1000m).ToString("0.###", CultureInfo.InvariantCulture) + "%"; changed = true; }
            }
            if (map.TryGetValue(element.Name.NamespaceName, out var ns)) { element.Name = XName.Get(element.Name.LocalName, ns); changed = true; }
            foreach (var attribute in element.Attributes().ToArray())
            {
                if (attribute.IsNamespaceDeclaration && map.TryGetValue(attribute.Value, out ns))
                { attribute.Value = ns; changed = true; }
                else if (!attribute.IsNamespaceDeclaration && map.TryGetValue(attribute.Name.NamespaceName, out ns))
                {
                    var name = XName.Get(attribute.Name.LocalName, ns); var value = attribute.Value;
                    attribute.Remove(); element.Add(new XAttribute(name, value)); changed = true;
                }
            }
            if (element.Name != PackageRel + "Relationship" || element.Attribute("Type") is not { } type) continue;
            var prefix = toStrict ? Rel.NamespaceName : StrictRel;
            if (!type.Value.StartsWith(prefix + "/", StringComparison.Ordinal)) continue;
            var suffix = type.Value[(prefix.Length + 1)..];
            suffix = toStrict ? suffix switch { "extended-properties" => "extendedProperties", "custom-properties" => "customProperties", _ => suffix }
                : suffix switch { "extendedProperties" => "extended-properties", "customProperties" => "custom-properties", _ => suffix };
            type.Value = (toStrict ? StrictRel : Rel.NamespaceName) + "/" + suffix; changed = true;
        }
        return changed;
    }
    internal static bool SourceIsStrict(WorkbookSnapshot snapshot, CancellationToken cancellationToken) =>
        snapshot.FileState is { } source && new XlsxPackage(source.OriginalWorkbook.ToArray(), new(cancellationToken)).IsStrict;

    internal static void PreserveVariant(WorkbookSnapshot snapshot, Dictionary<string, byte[]> parts, CancellationToken cancellationToken)
    {
        if (!SourceIsStrict(snapshot, cancellationToken)) return;
        var budget = new FileReadBudget(cancellationToken);
        foreach (var path in parts.Keys.Where(p => p.EndsWith(".xml", StringComparison.Ordinal) || p.EndsWith(".rels", StringComparison.Ordinal)).ToArray())
        {
            var document = Xml(parts[path], budget);
            if (document.Root?.Name.Namespace == XlsxThreadedComments.Thread) continue;
            var changed = Map(document, true, budget.Check);
            if (document.Root?.Name == XName.Get("workbook", StrictMain))
            { document.Root.SetAttributeValue("conformance", "strict"); changed = true; }
            if (changed) parts[path] = XmlBytes(document);
        }
    }

    internal static CellValue DateValue(string text, ExcelDateSystem system)
    {
        // Excel's native Strict writer uses a local calendar timestamp, including its synthetic 1900-02-29.
        var leapDay = text.StartsWith("1900-02-29", StringComparison.Ordinal);
        var parsed = leapDay ? "1900-02-28" + text[10..] : text;
        if (!DateTime.TryParseExact(parsed, ["yyyy-MM-dd", "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF"],
            CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) || leapDay && system == ExcelDateSystem.Mac1904)
            throw new NotSupportedException("日期型单元格需为 Excel 本地 ISO 日期/时间；时区及其他 ISO 形式尚未建模。");
        var origin = system == ExcelDateSystem.Mac1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31);
        var serial = (date - origin).TotalDays;
        if (system == ExcelDateSystem.Windows1900 && (leapDay || date >= new DateTime(1900, 3, 1))) serial++;
        return CellValue.FromNumber(serial);
    }
}
