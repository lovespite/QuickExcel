using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.TypedWorkbookXml;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static class TypedImages
{
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace Art = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static IReadOnlyList<SheetImage> Read(XlsxPackage package, string sheetPart, XElement sheet)
    {
        var images = new List<SheetImage>();
        foreach (var drawing in sheet.Elements(Main + "drawing"))
        {
            var path = package.Related(sheetPart, Attr(drawing, Rel + "id"), "drawing");
            package.RequireType(path, "application/vnd.openxmlformats-officedocument.drawing+xml");
            var root = package.XmlPart(path).Root!;
            if (root.Name != Drawing + "wsDr") throw new InvalidDataException("无效的绘图部件。");
            foreach (var anchor in root.Elements())
            {
                if (anchor.Name != Drawing + "twoCellAnchor" || anchor.Element(Drawing + "pic") is not { } pic)
                { package.Issues.Add($"{path}: 未建模绘图或锚点"); continue; }
                package.Shape(anchor, "from to pic clientData", "editAs", path);
                if ((string?)anchor.Attribute("editAs") is not (null or "twoCell")) package.Issues.Add($"{path}: 非随单元格缩放的图片锚点");
                int Marker(string name, string coordinate, int maximum)
                {
                    var value = anchor.Element(Drawing + name)?.Element(Drawing + coordinate)?.Value ?? throw new InvalidDataException("图片锚点缺失。");
                    return Int(value, maximum);
                }
                var row = Marker("from", "row", WorkbookLimits.MaxRows - 1);
                var column = Marker("from", "col", WorkbookLimits.MaxColumns - 1);
                var bottom = Marker("to", "row", WorkbookLimits.MaxRows);
                var right = Marker("to", "col", WorkbookLimits.MaxColumns);
                foreach (var name in new[] { "from", "to" })
                    if (Marker(name, "rowOff", int.MaxValue) != 0 || Marker(name, "colOff", int.MaxValue) != 0)
                        package.Issues.Add($"{path}: 图片像素偏移未建模");
                var blip = pic.Element(Drawing + "blipFill")?.Element(Art + "blip") ?? throw new InvalidDataException("图片引用缺失。");
                var imagePath = package.Related(path, Attr(blip, Rel + "embed"), "image");
                var extension = imagePath[(imagePath.LastIndexOf('.') + 1)..].ToLowerInvariant();
                package.RequireType(imagePath, Mime(extension));
                if (pic.Element(Drawing + "blipFill")?.Element(Art + "srcRect") is not null ||
                    pic.Element(Drawing + "spPr")?.Element(Art + "xfrm")?.HasAttributes == true)
                    package.Issues.Add($"{path}: 图片裁剪或旋转未建模");
                var imageBytes = package.BinaryPart(imagePath);
                package.Budget.Parse(imageBytes.Length + 256);
                images.Add(new SheetImage(imageBytes, extension, row, column, bottom - row, right - column));
            }
        }
        return images;
    }

    public static void Write(WorksheetSnapshot sheet, int sheetNumber, Dictionary<string, byte[]> parts, XElement root, XElement types)
    {
        if (sheet.Images.Length == 0) return;
        if (sheet.Images.Any(image => image.PlaceInCell)) throw new NotSupportedException("当前类型化写入器仅支持浮动图片；单元格内图片不能静默转换。");
        var path = $"xl/drawings/drawing{sheetNumber}.xml";
        var drawing = new XElement(Drawing + "wsDr", new XAttribute(XNamespace.Xmlns + "a", Art), new XAttribute(XNamespace.Xmlns + "r", Rel));
        var relationships = new XElement(PackageRel + "Relationships");
        for (var index = 0; index < sheet.Images.Length; index++)
        {
            var image = sheet.Images[index];
            var filename = $"image{sheetNumber}-{index + 1}.{image.Extension}";
            var mediaPath = "xl/media/" + filename;
            parts.Add(mediaPath, image.Data.ToArray());
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + mediaPath), new XAttribute("ContentType", Mime(image.Extension))));
            relationships.Add(Relationship($"image{index + 1}", "image", "../media/" + filename));
            XElement Marker(string name, int row, int col) => new(Drawing + name, new XElement(Drawing + "col", col), new XElement(Drawing + "colOff", 0), new XElement(Drawing + "row", row), new XElement(Drawing + "rowOff", 0));
            drawing.Add(new XElement(Drawing + "twoCellAnchor", Marker("from", image.RowIndex, image.ColumnIndex),
                Marker("to", image.RowIndex + image.RowSpan, image.ColumnIndex + image.ColumnSpan),
                new XElement(Drawing + "pic", new XElement(Drawing + "nvPicPr",
                    new XElement(Drawing + "cNvPr", new XAttribute("id", index + 1), new XAttribute("name", $"Image {index + 1}")), new XElement(Drawing + "cNvPicPr")),
                    new XElement(Drawing + "blipFill", new XElement(Art + "blip", new XAttribute(Rel + "embed", $"image{index + 1}")), new XElement(Art + "stretch", new XElement(Art + "fillRect"))),
                    new XElement(Drawing + "spPr", new XElement(Art + "prstGeom", new XAttribute("prst", "rect"), new XElement(Art + "avLst")))), new XElement(Drawing + "clientData")));
        }
        root.Add(new XElement(Main + "drawing", new XAttribute(Rel + "id", "drawing")));
        parts[RelationshipPart($"xl/worksheets/sheet{sheetNumber}.xml")] = XmlBytes(new XDocument(new XElement(PackageRel + "Relationships", Relationship("drawing", "drawing", $"../drawings/drawing{sheetNumber}.xml"))));
        parts[RelationshipPart(path)] = XmlBytes(new XDocument(relationships));
        parts[path] = XmlBytes(new XDocument(drawing));
        types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
    }
    private static string Mime(string extension) => extension switch
    {
        "png" => "image/png", "jpg" or "jpeg" => "image/jpeg", "gif" => "image/gif", "bmp" => "image/bmp", "tif" or "tiff" => "image/tiff",
        _ => throw new NotSupportedException("图片格式不支持。")
    };
}
