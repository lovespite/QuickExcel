using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace QuickExcel.CompatibilityProbe;

/// <summary>The SDK translates Strict names but rejects valid percent lexical values. Validate an explicit temporary projection, preserving all raw package facts.</summary>
internal static class StrictAuditProjection
{
    internal static byte[] ForSdk(Dictionary<string, byte[]> parts, Dictionary<string, XDocument> xml, SortedDictionary<string, string> facts, byte[] original)
    {
        XNamespace drawing = "http://purl.oclc.org/ooxml/drawingml/main";
        if (!xml.Values.Any(d => d.Root?.Name == XName.Get("workbook", "http://purl.oclc.org/ooxml/spreadsheetml/main"))) return original;
        facts["validation|profile"] = "Strict input; Open XML SDK namespace translation plus exact DrawingML percentage-to-integer validation projection. This is not a complete ISO Strict schema validator; native Excel verification is separate.";
        var projected = new Dictionary<string, byte[]>(parts, StringComparer.Ordinal); var count = 0;
        var percentages = "tint shade alpha alphaOff alphaMod hueMod sat satOff satMod lum lumOff lumMod red redOff redMod green greenOff greenMod blue blueOff blueMod".Split(' ').ToHashSet(StringComparer.Ordinal);
        foreach (var (part, source) in xml)
        {
            var copy = new XDocument(source); var changed = false;
            foreach (var e in copy.Descendants().Where(e => e.Name.Namespace == drawing))
            foreach (var a in e.Attributes().Where(a => a.Name.NamespaceName.Length == 0 && a.Value.EndsWith('%') &&
                (a.Name.LocalName == "val" && percentages.Contains(e.Name.LocalName) || a.Name.LocalName == "pos" && e.Name.LocalName == "gs" ||
                a.Name.LocalName == "lim" && e.Name.LocalName == "miter" ||
                a.Name.LocalName is "l" or "t" or "r" or "b" && e.Name.LocalName is "srcRect" or "fillRect" or "tileRect")))
            {
                if (!decimal.TryParse(a.Value[..^1], NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var value) ||
                    value < int.MinValue / 1000m || value > int.MaxValue / 1000m || value * 1000 != decimal.Truncate(value * 1000)) continue;
                a.Value = (value * 1000).ToString("0", CultureInfo.InvariantCulture); changed = true; count++;
            }
            if (changed) projected[part] = Encoding.UTF8.GetBytes(copy.ToString(SaveOptions.DisableFormatting));
        }
        facts["validation|projected-percentages"] = count.ToString(CultureInfo.InvariantCulture);
        return count == 0 ? original : OpenXmlAudit.Zip(projected);
    }
}
