using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace QuickExcel.CompatibilityProbe;

public sealed record AuditIssue(string Part, string Location, string Message);
public sealed record PackageAudit(string Sha256, SortedDictionary<string, string> Facts, IReadOnlyList<AuditIssue> Issues);
public sealed record PackageDifference(string Category, string Location, string? Before, string? After);

/// <summary>Development-only independent ZIP/XML and SDK inspection. Does not use the QuickExcel reader.</summary>
public static class OpenXmlAudit
{
    public static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const long Limit = 64 * 1024 * 1024;

    public static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    public static Dictionary<string, byte[]> ReadParts(byte[] bytes)
    {
        if (bytes.LongLength > Limit) throw new InvalidDataException("Probe input exceeds 64 MiB.");
        using var stream = new MemoryStream(bytes, false);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        if (zip.Entries.Count > 2048) throw new InvalidDataException("Probe entry count exceeds 2048.");
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long total = 0;
        foreach (var entry in zip.Entries)
        {
            var name = entry.FullName;
            if (name.Length == 0 || name.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 ||
                name.Split('/').Any(s => s is "" or "." or "..") || !names.Add(name))
                throw new InvalidDataException($"Ambiguous package entry: {name}");
            if (entry.Length > Limit - total) throw new InvalidDataException("Probe expanded package exceeds 64 MiB.");
            using var input = entry.Open();
            using var output = new MemoryStream();
            var buffer = new byte[8192];
            int count;
            while ((count = input.Read(buffer)) != 0)
            {
                total += count;
                if (total > Limit) throw new InvalidDataException("Probe expanded package exceeds 64 MiB.");
                output.Write(buffer, 0, count);
            }
            parts.Add(name, output.ToArray());
        }
        return parts;
    }

    public static byte[] Zip(IReadOnlyDictionary<string, byte[]> parts)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var (name, bytes) in parts.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                var entry = zip.CreateEntry(name);
                entry.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var target = entry.Open();
                target.Write(bytes);
            }
        return output.ToArray();
    }

    public static XDocument Xml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, false);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = Limit
        });
        var doc = XDocument.Load(reader);
        if (doc.Root is null || doc.Descendants().Any(e => e.Ancestors().Take(65).Count() > 64))
            throw new InvalidDataException("Probe XML depth exceeds 64.");
        return doc;
    }

    public static PackageAudit Inspect(byte[] bytes)
    {
        var facts = new SortedDictionary<string, string>(StringComparer.Ordinal);
        var issues = new List<AuditIssue>();
        var parts = ReadParts(bytes);
        var xml = new Dictionary<string, XDocument>(StringComparer.Ordinal);
        foreach (var (part, data) in parts)
        {
            facts[$"parts|{part}"] = Hash(data);
            if (!part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && !part.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            try { xml.Add(part, Xml(data)); }
            catch (Exception ex) when (ex is XmlException or InvalidDataException)
            { issues.Add(new(part, "/", ex.Message)); }
        }
        foreach (var (part, document) in xml)
        {
            Flatten(document.Root!, part, "", facts);
            if (part.EndsWith(".rels", StringComparison.Ordinal))
            {
                var source = RelationshipSource(part);
                if (source != "" && !parts.ContainsKey(source)) issues.Add(new(part, source, "Relationship source is missing."));
                foreach (var relationship in document.Root!.Elements(PackageRel + "Relationship"))
                {
                    if ((string?)relationship.Attribute("TargetMode") == "External") continue; // Never dereference external links.
                    try
                    {
                        var target = Resolve(source, (string?)relationship.Attribute("Target") ?? "");
                        if (!parts.ContainsKey(target)) issues.Add(new(part, (string?)relationship.Attribute("Id") ?? "", $"Relationship target is missing: {target}"));
                    }
                    catch (InvalidDataException ex) { issues.Add(new(part, "/", ex.Message)); }
                }
            }
            else
            {
                var relPart = RelationshipPart(part);
                var ids = xml.GetValueOrDefault(relPart)?.Root?.Elements(PackageRel + "Relationship")
                    .Select(r => (string?)r.Attribute("Id")).ToHashSet() ?? [];
                foreach (var reference in document.Descendants().Attributes().Where(a => a.Name.Namespace == Rel))
                    if (!ids.Contains(reference.Value)) issues.Add(new(part, reference.Parent!.Name.LocalName, $"Unresolved relationship: {reference.Value}"));
            }
        }
        AddCells(xml, facts, issues);
        try
        {
            using var stream = new MemoryStream(bytes, false);
            using var document = SpreadsheetDocument.Open(stream, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019) { MaxNumberOfErrors = 1000 };
            issues.AddRange(validator.Validate(document).Select(e => new AuditIssue(
                e.Part?.Uri.ToString() ?? "package", e.Path?.XPath ?? "", e.Description)));
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or XmlException or ArgumentException or InvalidOperationException)
        { issues.Add(new("package", "/", ex.Message)); }
        return new(Hash(bytes), facts, issues);
    }

    public static IReadOnlyList<PackageDifference> Compare(PackageAudit before, PackageAudit after) =>
        before.Facts.Keys.Union(after.Facts.Keys).Order(StringComparer.Ordinal)
            .Where(key => before.Facts.GetValueOrDefault(key) != after.Facts.GetValueOrDefault(key))
            .Select(key => new PackageDifference(key[..key.IndexOf('|')], key[(key.IndexOf('|') + 1)..],
                before.Facts.GetValueOrDefault(key), after.Facts.GetValueOrDefault(key))).ToArray();

    private static void Flatten(XElement element, string part, string parent, IDictionary<string, string> facts)
    {
        // Expanded names prevent namespace-prefix changes from being reported as semantic changes.
        var identity = element.Attribute("r") ?? element.Attribute("Id") ?? element.Attribute("name") ??
            element.Attribute("PartName") ?? element.Attribute("Extension");
        var index = element.ElementsBeforeSelf(element.Name).Count() + 1;
        var path = parent + "/" + element.Name + (identity is null ? $"[{index}]" : $"[{identity.Name}={JsonSerializer.Serialize(identity.Value)};{index}]");
        var category = part.EndsWith(".rels", StringComparison.Ordinal) ? "relationships" :
            element.AncestorsAndSelf(Main + "styleSheet").Any() ? "styles" :
            element.AncestorsAndSelf(Main + "f").Any() ? "formulas" :
            element.AncestorsAndSelf(Main + "c").Any() ? "content" : "metadata";
        facts[$"{category}|{part}{path}"] = element.HasElements ? "element" : element.Value;
        foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            facts[$"{category}|{part}{path}/@{attribute.Name}"] = attribute.Value;
        foreach (var child in element.Elements()) Flatten(child, part, path, facts);
    }

    private static void AddCells(Dictionary<string, XDocument> xml, IDictionary<string, string> facts, List<AuditIssue> issues)
    {
        foreach (var (part, document) in xml.Where(p => p.Value.Root?.Name == Main + "worksheet"))
        {
            // Resolve styles and strings through the owning workbook's relationships, never fixed filenames.
            var workbook = xml.FirstOrDefault(p => p.Value.Root?.Name == Main + "workbook" && Related(xml, p.Key, "worksheet").Contains(part)).Key;
            if (workbook is null) continue;
            var stylePart = Related(xml, workbook, "styles").FirstOrDefault();
            var stringPart = Related(xml, workbook, "sharedStrings").FirstOrDefault();
            var xfs = stylePart is null ? [] : xml.GetValueOrDefault(stylePart)?.Root?.Element(Main + "cellXfs")?.Elements().ToArray() ?? [];
            var strings = stringPart is null ? [] : xml.GetValueOrDefault(stringPart)?.Root?.Elements(Main + "si").ToArray() ?? [];
            foreach (var cell in document.Descendants(Main + "c"))
            {
                var address = (string?)cell.Attribute("r") ?? "(implicit)";
                var type = (string?)cell.Attribute("t") ?? "n";
                var value = cell.Element(Main + "v")?.Value;
                if (type == "inlineStr") { type = "text"; value = string.Concat(cell.Element(Main + "is")?.Descendants(Main + "t").Select(t => t.Value) ?? []); }
                if (type == "s")
                {
                    type = "text";
                    if (int.TryParse(value, out var index) && index >= 0 && index < strings.Length)
                        value = string.Concat(strings[index].Descendants(Main + "t").Select(t => t.Value));
                    else issues.Add(new(part, address, "Shared string index is outside the string table."));
                }
                if (type == "n" && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    value = number.ToString("R", CultureInfo.InvariantCulture);
                facts[$"cells|{part}!{address}/value"] = JsonSerializer.Serialize(new { type, value });
                if (cell.Element(Main + "f") is { } formula) facts[$"cells|{part}!{address}/formula"] = formula.Value;
                if (cell.Attribute("s") is { } style && (!int.TryParse(style.Value, out var id) || id < 0 || id >= xfs.Length))
                    issues.Add(new(part, address, "Style index is outside cellXfs."));
            }
        }
    }

    private static IEnumerable<string> Related(Dictionary<string, XDocument> xml, string source, string type) =>
        xml.GetValueOrDefault(RelationshipPart(source))?.Root?.Elements(PackageRel + "Relationship")
            .Where(r => (string?)r.Attribute("Type") == Rel.NamespaceName + "/" + type && (string?)r.Attribute("TargetMode") != "External")
            .Select(r => Resolve(source, (string?)r.Attribute("Target") ?? "")) ?? [];

    private static string RelationshipPart(string part)
    {
        var slash = part.LastIndexOf('/');
        return part[..(slash + 1)] + "_rels/" + part[(slash + 1)..] + ".rels";
    }

    private static string RelationshipSource(string part)
    {
        if (part == "_rels/.rels") return "";
        var marker = part.LastIndexOf("_rels/", StringComparison.Ordinal);
        if (marker < 0) throw new InvalidDataException("Invalid relationship part path.");
        return part[..marker] + part[(marker + 6)..^5];
    }

    private static string Resolve(string source, string target)
    {
        if (target.Length == 0 || target.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0)
            throw new InvalidDataException("Invalid internal relationship target.");
        var segments = target.StartsWith('/') ? new List<string>() : source.Split('/').SkipLast(1).ToList();
        foreach (var segment in target.TrimStart('/').Split('/'))
        {
            if (segment == "..")
            {
                if (segments.Count == 0) throw new InvalidDataException("Relationship escapes package root.");
                segments.RemoveAt(segments.Count - 1);
            }
            else if (segment is not ("" or ".")) segments.Add(segment);
        }
        return string.Join('/', segments);
    }
}
