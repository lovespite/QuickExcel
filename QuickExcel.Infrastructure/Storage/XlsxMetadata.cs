using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;

namespace QuickExcel.Storage;

/// <summary>Dependency-free document properties, plus a verified worksheet title index.
/// Unknown elements, linked custom properties and outbound relationships never enter the whitelist.</summary>
internal static class XlsxMetadata
{
    private static readonly XNamespace Core = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Terms = "http://purl.org/dc/terms/";
    private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";
    private static readonly XNamespace App = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace Custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    private sealed record Spec(string Name, string Relationship, string Type, XName Root);
    private sealed record PropertyPart(Spec Spec, string Path, XElement Root, byte[] Bytes);
    private static readonly Spec[] Specs =
    [
        new("core", PackageRel.NamespaceName + "/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml", Core + "coreProperties"),
        new("app", Rel.NamespaceName + "/extended-properties", "application/vnd.openxmlformats-officedocument.extended-properties+xml", App + "Properties"),
        new("custom", Rel.NamespaceName + "/custom-properties", "application/vnd.openxmlformats-officedocument.custom-properties+xml", Custom + "Properties")
    ];

    public static void Inspect(XlsxPackage package, IReadOnlyList<string> sheetNames) => _ = Read(package, sheetNames);

    private static List<PropertyPart> Read(XlsxPackage package, IReadOnlyList<string> sheetNames)
    {
        var result = new List<PropertyPart>();
        foreach (var spec in Specs)
        {
            if (package.RelatedByUri("", spec.Relationship) is not { } path) continue;
            package.RequireType(path, spec.Type);
            var root = package.XmlPart(path).Root!;
            if (root.Name != spec.Root) throw new InvalidDataException(path + ": 文档属性根元素无效。");
            var safe = package.Relationships(path).Count == 0 && Attributes(root) && spec.Name switch
            {
                "core" => CoreProperties(root),
                "custom" => CustomProperties(root),
                _ => AppProperties(root, sheetNames)
            };
            if (safe) result.Add(new(spec, path, root, package.BinaryPart(path)));
            else package.Loss(path + ": 文档属性含白名单以外的结构或依赖，编辑后不能可靠保留。", "metadata/" + spec.Name, path);
        }
        return result;
    }

    private static bool Attributes(XElement e, params XName[] allowed) =>
        e.Attributes().All(a => a.IsNamespaceDeclaration || allowed.Contains(a.Name));
    private static bool Scalar(XElement e) => !e.HasElements && Attributes(e);
    private static bool UniqueChildren(XElement root) => root.Elements().Select(e => e.Name).Distinct().Count() == root.Elements().Count();

    private static bool CoreProperties(XElement root) => UniqueChildren(root) && root.Elements().All(e =>
    {
        if (e.Name == Terms + "created" || e.Name == Terms + "modified")
        {
            var type = (string?)e.Attribute(Xsi + "type");
            var parts = type?.Split(':');
            return !e.HasElements && Attributes(e, Xsi + "type") && parts is { Length: 2 } && parts[1] == "W3CDTF" &&
                e.GetNamespaceOfPrefix(parts[0]) == Terms && Date(e.Value);
        }
        if (!Scalar(e)) return false;
        if (e.Name == Core + "lastPrinted") return Date(e.Value);
        return e.Name.Namespace == Dc && e.Name.LocalName is "title" or "subject" or "creator" or "description" or "identifier" or "language" ||
            e.Name.Namespace == Core && e.Name.LocalName is "keywords" or "lastModifiedBy" or "revision" or "category" or "contentStatus" or "version";
    });

    private static bool Date(string value)
    {
        try { _ = XmlConvert.ToDateTimeOffset(value); return true; }
        catch (FormatException) { return false; }
    }
    private static bool Boolean(string value) => value is "true" or "false" or "0" or "1";
    private static bool CustomProperties(XElement root)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var ids = new HashSet<int>();
        return root.Elements().All(e => e.Name == Custom + "property" && Attributes(e, "fmtid", "pid", "name") &&
            Guid.TryParse((string?)e.Attribute("fmtid"), out var format) && format == new Guid("D5CDD505-2E9C-101B-9397-08002B2CF9AE") &&
            int.TryParse((string?)e.Attribute("pid"), NumberStyles.None, CultureInfo.InvariantCulture, out var id) && id >= 2 && ids.Add(id) &&
            (string?)e.Attribute("name") is { Length: > 0 } name && names.Add(name) && e.Elements().Count() == 1 && Variant(e.Elements().Single()));
    }
    private static bool Variant(XElement e)
    {
        if (e.Name.Namespace != Vt || !Scalar(e)) return false;
        return e.Name.LocalName switch
        {
            "lpwstr" or "lpstr" or "bstr" => true,
            "bool" => Boolean(e.Value),
            "i4" or "int" => int.TryParse(e.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _),
            "i8" => long.TryParse(e.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _),
            "r8" => double.TryParse(e.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number),
            "filetime" or "date" => Date(e.Value),
            _ => false
        };
    }
    private static bool AppProperties(XElement root, IReadOnlyList<string> sheetNames)
    {
        if (!UniqueChildren(root)) return false;
        foreach (var e in root.Elements())
        {
            if (e.Name.Namespace != App) return false;
            if (e.Name.LocalName is "HeadingPairs" or "TitlesOfParts") continue;
            if (!Scalar(e)) return false;
            if (e.Name.LocalName is "Application" or "AppVersion" or "Company" or "Manager" or "Template" or "HyperlinkBase") continue;
            if (e.Name.LocalName is "ScaleCrop" or "LinksUpToDate" or "SharedDoc" or "HyperlinksChanged")
            { if (!Boolean(e.Value)) return false; continue; }
            if (e.Name.LocalName == "DocSecurity" && uint.TryParse(e.Value, NumberStyles.None, CultureInfo.InvariantCulture, out _)) continue;
            return false;
        }
        var headings = root.Element(App + "HeadingPairs"); var titles = root.Element(App + "TitlesOfParts");
        if (headings is null && titles is null) return true;
        if (headings is null || titles is null || !Attributes(headings) || !Attributes(titles)) return false;
        if (headings.Elements().Count() != 1 || titles.Elements().Count() != 1) return false;
        var h = headings.Elements().Single(); var t = titles.Elements().Single();
        if (h.Name != Vt + "vector" || !Attributes(h, "size", "baseType") || (string?)h.Attribute("baseType") != "variant" ||
            (string?)h.Attribute("size") != "2" || h.Elements().Count() != 2 ||
            h.Elements().Any(e => e.Name != Vt + "variant" || !Attributes(e) || e.Elements().Count() != 1)) return false;
        var label = h.Elements().First().Elements().Single(); var count = h.Elements().Last().Elements().Single();
        // Other groups/localizations require a separately verified mapping, even if their titles happen to match.
        return label.Name == Vt + "lpstr" && Scalar(label) && label.Value is "Worksheets" or "工作表" && count.Name == Vt + "i4" && Scalar(count) &&
            count.Value == sheetNames.Count.ToString(CultureInfo.InvariantCulture) && t.Name == Vt + "vector" && Attributes(t, "size", "baseType") &&
            (string?)t.Attribute("baseType") == "lpstr" && (string?)t.Attribute("size") == count.Value &&
            t.Elements().All(e => e.Name == Vt + "lpstr" && Scalar(e)) &&
            t.Elements().Select(e => Decode(e.Value)).SequenceEqual(sheetNames);
    }

    private static List<PropertyPart> Source(WorkbookSnapshot snapshot, CancellationToken cancellationToken = default) => snapshot.FileState is { } state
        ? Read(new XlsxPackage(state.OriginalWorkbook.ToArray(), new(cancellationToken)), state.OriginalModel.Sheets.Select(s => s.Name).ToArray()) : [];
    private static bool TitlesChanged(WorkbookSnapshot snapshot) => snapshot.FileState is { } state &&
        !snapshot.Sheets.Select(s => s.Name).SequenceEqual(state.OriginalModel.Sheets.Select(s => s.Name));

    public static IEnumerable<WorkbookSavePart> Plan(WorkbookSnapshot snapshot) => Source(snapshot).Select(p =>
        new WorkbookSavePart(p.Path, p.Spec.Name == "app" && p.Root.Element(App + "TitlesOfParts") is not null && TitlesChanged(snapshot) ? "update-sheet-titles" : "preserve",
            p.Spec.Name == "app" ? "已验证的静态属性；工作表标题索引依赖名称和顺序。" : "已验证的无关系标量属性，不依赖单元格与工作表。"));

    public static void Write(WorkbookSnapshot snapshot, Dictionary<string, byte[]> parts, XElement relationships, XElement types, CancellationToken cancellationToken)
    {
        foreach (var part in Source(snapshot, cancellationToken))
        {
            var path = "docProps/" + part.Spec.Name + ".xml";
            var bytes = part.Bytes;
            if (part.Spec.Name == "app" && part.Root.Element(App + "TitlesOfParts") is not null && TitlesChanged(snapshot))
            {
                var root = new XElement(part.Root);
                root.Element(App + "HeadingPairs")!.Descendants(Vt + "i4").Single().Value = snapshot.Sheets.Length.ToString(CultureInfo.InvariantCulture);
                var titles = root.Element(App + "TitlesOfParts")!.Element(Vt + "vector")!;
                titles.SetAttributeValue("size", snapshot.Sheets.Length);
                titles.ReplaceNodes(snapshot.Sheets.Select(s => new XElement(Vt + "lpstr", Encode(s.Name))));
                bytes = XmlBytes(new XDocument(root));
            }
            parts.Add(path, bytes);
            relationships.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "props-" + part.Spec.Name),
                new XAttribute("Type", part.Spec.Relationship), new XAttribute("Target", path)));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", part.Spec.Type)));
        }
    }

    public static void VerifyWritten(WorkbookSnapshot snapshot, byte[] bytes, CancellationToken cancellationToken)
    {
        var expected = new Dictionary<string, byte[]>();
        Write(snapshot, expected, new XElement(PackageRel + "Relationships"), new XElement(ContentTypes + "Types"), cancellationToken);
        if (expected.Count == 0) return;
        var written = new XlsxPackage(bytes, new(cancellationToken));
        var recognized = Read(written, snapshot.Sheets.Select(s => s.Name).ToArray());
        if (recognized.Count != expected.Count) throw new InvalidDataException("保存后文档属性关系或结构校验不一致。");
        foreach (var (path, content) in expected)
            if (!written.Parts.TryGetValue(path, out var saved) || !content.SequenceEqual(saved))
                throw new InvalidDataException(path + ": 保存后的属性与保留计划不一致。");
    }
}
