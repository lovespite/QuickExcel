using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    // An unchanged workbook normally uses the original-package path. Repair our known
    // ChartEx defects there too, without rebuilding sheets or dropping opaque content.
    internal static byte[] RepairPreservedExtendedCharts(byte[] original, WorkbookSnapshot snapshot, CancellationToken cancellationToken)
    {
        var charts = snapshot.Sheets.SelectMany(s => s.Data.Charts).Where(c => c.IsExtended && !c.OriginalXml.IsEmpty).ToArray();
        if (charts.Length == 0) return original;
        var budget = new FileReadBudget(cancellationToken);
        var package = new XlsxPackage(original, budget);
        var parts = package.Parts;
        var types = Xml(parts["[Content_Types].xml"], budget).Root!;
        var changed = false;
        foreach (var item in types.Elements(ContentTypes + "Override").Where(e => (string?)e.Attribute("ContentType") == ChartExType).ToArray())
        {
            budget.Check();
            var path = ((string)item.Attribute("PartName")!)[1..];
            if (!parts.TryGetValue(path, out var raw)) continue;
            // Only charts successfully read by the closed schema are eligible. Never
            // rewrite an unknown ChartEx layout/extension from an imported package.
            var candidates = charts.Where(c => c.OriginalXml.AsSpan().SequenceEqual(raw)).ToArray();
            if (candidates.Length == 0) continue;
            System.Collections.Immutable.ImmutableArray<ChartStylePart> importedStyles;
            try { importedStyles = ReadStyleParts(package, path); }
            catch (NotSupportedException) { continue; }
            // Identical chart XML can refer to different style parts. Match those too,
            // including incomplete legacy styles normalized during import.
            var chart = candidates.FirstOrDefault(c => ExtendedStyleParts(c with { StyleParts = importedStyles })
                .Where(s => importedStyles.Any(p => p.Kind == s.Kind))
                .All(s => c.StyleParts.Any(p => p.Kind == s.Kind && p.Bytes.SequenceEqual(s.Bytes))));
            if (chart is null) continue;
            var xml = Xml(raw, budget);
            var axes = xml.Descendants(ChartEx + "axisId").ToArray();
            if (axes.Length > 0)
            {
                axes.Remove(); // Includes the full-series copy in our filter extension.
                parts[path] = XmlBytes(xml); changed = true;
            }
            var relPath = RelationshipPart(path);
            var rels = parts.TryGetValue(relPath, out var relationships) ? Xml(relationships, budget).Root! : new XElement(PackageRel + "Relationships");
            var added = false;
            foreach (var style in ExtendedStyleParts(chart))
            {
                if (rels.Elements().FirstOrDefault(e => (string?)e.Attribute("Type") == StyleRelationship + style.Kind) is { } relationship)
                {
                    var existingPath = Resolve(path, (string)relationship.Attribute("Target")!);
                    if (!style.Bytes.AsSpan().SequenceEqual(parts[existingPath])) { parts[existingPath] = style.Bytes.ToArray(); changed = true; }
                    continue;
                }
                var id = style.RelationshipId;
                while (rels.Elements().Any(e => (string?)e.Attribute("Id") == id)) id += "1";
                var stem = path[..path.LastIndexOf('.')] + "-excel-" + style.Kind;
                var target = stem + ".xml";
                for (var suffix = 1; parts.Keys.Any(p => string.Equals(p, target, StringComparison.OrdinalIgnoreCase)); suffix++) target = stem + suffix + ".xml";
                parts.Add(target, style.Bytes.ToArray());
                rels.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", StyleRelationship + style.Kind), new XAttribute("Target", Path.GetFileName(target))));
                types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + target), new XAttribute("ContentType", StyleContentType(style.Kind))));
                added = changed = true;
            }
            if (added) parts[relPath] = XmlBytes(new XDocument(rels));
        }
        if (!changed) return original;
        parts["[Content_Types].xml"] = XmlBytes(new XDocument(types));
        budget.Check();
        return Zip(parts);
    }
}
