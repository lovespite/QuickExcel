using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;
using static QuickExcel.Storage.TypedWorkbookXml;

namespace QuickExcel.Storage;

internal static class XlsxAnnotations
{
    private const string CommentsType = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    private const string VmlType = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml", Excel = "urn:schemas-microsoft-com:office:excel";
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
    public static WorksheetAnnotations Read(XlsxPackage package, string path, XElement sheet, XlsxCellText textReader)
    {
        var links = new Dictionary<CellPosition, CellHyperlink>(); var notes = new Dictionary<CellPosition, CellNote>(); byte[] drawing = [];
        if (sheet.Elements(Main + "hyperlinks").Count() > 1 || sheet.Elements(Main + "legacyDrawing").Count() > 1) throw new InvalidDataException("重复链接/批注绘图定义。");
        if (sheet.Element(Main + "hyperlinks") is { } list)
        {
            package.Shape(list, "hyperlink", "", path);
            foreach (var element in list.Elements(Main + "hyperlink"))
            {
                package.Budget.Parse(256);
                if (links.Count >= 65530) throw new InvalidDataException("超链接数量过多。");
                var shape = new XElement(element); shape.Attribute(Rel + "id")?.Remove(); package.Shape(shape, "", "ref location display tooltip", path);
                var reference = Attr(element, "ref");
                if (reference.Contains(':')) { package.Loss($"{path}: 多格超链接区域尚未建模", "hyperlink/range", path, element); continue; }
                var address = Position(reference);
                var target = element.Attribute(Rel + "id") is { } id ? package.ExternalHyperlink(path, id.Value) : null;
                string? Text(string name) => element.Attribute(name) is { } a ? Decode(a.Value) : null;
                if (!links.TryAdd(address, new(target, Text("location"), Text("display"), Text("tooltip")))) throw new InvalidDataException("单元格超链接重复。");
            }
        }
        if (package.RelatedByType(path, "comments") is { } comments)
        {
            package.RequireType(comments, CommentsType); var root = package.XmlPart(comments).Root!;
            if (root.Name != Main + "comments") throw new InvalidDataException("批注部件无效。");
            package.Shape(root, "authors commentList", "", comments);
            if (root.Elements(Main + "authors").Count() != 1 || root.Elements(Main + "commentList").Count() != 1) throw new InvalidDataException("批注作者或列表定义无效。");
            package.Shape(root.Element(Main + "authors")!, "author", "", comments);
            foreach (var author in root.Element(Main + "authors")!.Elements(Main + "author")) package.Shape(author, "", "", comments);
            package.Shape(root.Element(Main + "commentList")!, "comment", "", comments);
            var authors = root.Element(Main + "authors")?.Elements(Main + "author").Select(a => Decode(a.Value)).ToArray() ?? [];
            if (authors.Length > 65530) throw new InvalidDataException("批注作者数量过多。");
            foreach (var comment in root.Element(Main + "commentList")?.Elements(Main + "comment") ?? [])
            {
                package.Budget.Parse(256);
                if (notes.Count >= 65530) throw new InvalidDataException("批注数量过多。");
                package.Shape(comment, "text", "ref authorId", comments);
                var address = Position(Attr(comment, "ref")); var author = authors[Int(Attr(comment, "authorId"), authors.Length - 1)];
                if (comment.Elements(Main + "text").Count() != 1 || comment.Element(Main + "text") is not { } text) throw new InvalidDataException("批注文本定义无效。");
                var parsed = textReader.Read(text, comments);
                if (!notes.TryAdd(address, new(author, parsed.Rich ?? new([new(parsed.Text)])))) throw new InvalidDataException("单元格批注重复。");
            }
        }
        if (sheet.Element(Main + "legacyDrawing") is { } legacy)
        {
            var part = package.Related(path, Attr(legacy, Rel + "id"), "vmlDrawing");
            package.RequireType(part, VmlType);
            var xml = package.XmlPart(part); var bytes = package.BinaryPart(part);
            if (bytes.Length > 1024 * 1024) throw new InvalidDataException("批注 VML 超过 1 MiB。");
            // Note-only VML has no workbook formulas or outgoing resources. Preserve geometry/visibility byte-for-byte.
            var shapes = xml.Descendants(Vml + "shape").ToArray();
            var allowed = new HashSet<string>(["MoveWithCells", "SizeWithCells", "Anchor", "AutoFill", "Row", "Column", "Visible"]);
            var allowedElements = new HashSet<XName>(new[] { "shape", "shapetype", "stroke", "path", "fill", "shadow", "textbox" }.Select(n => Vml + n)
                .Concat(allowed.Append("ClientData").Select(n => Excel + n)).Concat([Office + "shapelayout", Office + "idmap", XName.Get("xml"), XName.Get("div")]));
            var valid = package.Relationships(part).Count == 0 && shapes.Length == notes.Count && xml.Descendants(Excel + "ClientData").Count() == shapes.Length &&
                !xml.Descendants().Any(e => !allowedElements.Contains(e.Name)) &&
                !xml.Descendants().Any(e => e.Name.LocalName is "imagedata" or "formulas" or "OLEObject") &&
                !xml.Descendants().Attributes().Any(a => a.Name == Rel + "id" || a.Name.LocalName is "href" or "src" or "relid");
            var seen = new HashSet<CellPosition>();
            foreach (var shape in shapes)
            {
                var data = shape.Element(Excel + "ClientData");
                if (shape.Elements(Excel + "ClientData").Count() != 1 || (string?)data?.Attribute("ObjectType") != "Note" || data!.Elements().Any(e => e.Name.Namespace != Excel || !allowed.Contains(e.Name.LocalName))) { valid = false; continue; }
                if (!int.TryParse(data.Element(Excel + "Row")?.Value, out var row) || !int.TryParse(data.Element(Excel + "Column")?.Value, out var col) ||
                    !notes.ContainsKey(new(row, col)) || !seen.Add(new(row, col))) valid = false;
            }
            if (valid) drawing = bytes;
            else package.Loss($"{part}: VML 不属于无外部依赖的普通批注绘图", "comments/vml", part);
        }
        return new(links, notes, drawing);
    }

    public static void WriteLinks(WorksheetAnnotations annotations, XElement root)
    {
        if (annotations.Links.IsEmpty) return;
        root.Add(new XElement(Main + "hyperlinks", annotations.Links.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column).Select((p, i) =>
            new XElement(Main + "hyperlink", new XAttribute("ref", Reference(p.Key)),
                p.Value.Target is null ? null : new XAttribute(Rel + "id", "link" + i),
                p.Value.Location is null ? null : new XAttribute("location", Encode(p.Value.Location)),
                p.Value.Display is null ? null : new XAttribute("display", Encode(p.Value.Display)),
                p.Value.Tooltip is null ? null : new XAttribute("tooltip", Encode(p.Value.Tooltip))))));
    }

    public static void Write(WorksheetAnnotations annotations, int index, Dictionary<string, byte[]> parts, XElement root, XElement types)
    {
        if (annotations.IsEmpty) return;
        var relationsPath = $"xl/worksheets/_rels/sheet{index}.xml.rels";
        var rels = parts.TryGetValue(relationsPath, out var bytes) ? Xml(bytes, new(CancellationToken.None)).Root! : new XElement(PackageRel + "Relationships");
        void AddPart(string name, string type, byte[] content)
        { parts[name] = content; types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + name), new XAttribute("ContentType", type))); }
        var i = 0;
        foreach (var link in annotations.Links.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
        {
            if (link.Value.Target is { } target)
            { var rel = Relationship("link" + i, "hyperlink", target); rel.SetAttributeValue("TargetMode", "External"); rels.Add(rel); }
            i++;
        }
        if (!annotations.Notes.IsEmpty)
        {
            var authors = annotations.Notes.Values.Select(n => n.Author).Distinct().ToArray();
            var comments = new XElement(Main + "comments", new XElement(Main + "authors", authors.Select(a => new XElement(Main + "author", Encode(a)))),
                new XElement(Main + "commentList", annotations.Notes.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column).Select(p =>
                    new XElement(Main + "comment", new XAttribute("ref", Reference(p.Key)), new XAttribute("authorId", Array.IndexOf(authors, p.Value.Author)), XlsxCellText.Write(Main + "text", p.Value.Text)))));
            AddPart($"xl/comments{index}.xml", CommentsType, XmlBytes(new XDocument(comments)));
            rels.Add(Relationship("comments", "comments", $"../comments{index}.xml"));
        }
        if (!annotations.NoteDrawing.IsEmpty)
        {
            AddPart($"xl/drawings/notes{index}.vml", VmlType, annotations.NoteDrawing.ToArray());
            rels.Add(Relationship("notes", "vmlDrawing", $"../drawings/notes{index}.vml"));
            root.Add(new XElement(Main + "legacyDrawing", new XAttribute(Rel + "id", "notes")));
        }
        if (rels.HasElements) parts[relationsPath] = XmlBytes(new XDocument(rels));
    }
}
