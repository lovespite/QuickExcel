using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

internal static class XlsxDrawingAnchors
{
    internal static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    internal static readonly XNamespace Art = "http://schemas.openxmlformats.org/drawingml/2006/main";
    internal static DrawingAnchor Read(XElement element, WorksheetSnapshot geometry)
    {
        if (element.Name.Namespace != Drawing || element.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "editAs") ||
            element.Elements().Any(e => e.Name.Namespace != Drawing || e.Name.LocalName is not ("from" or "to" or "pos" or "ext" or "pic" or "graphicFrame" or "clientData")) ||
            element.Elements(Drawing + "clientData").Any(e => e.HasAttributes || e.HasElements)) throw new NotSupportedException("绘图锚点附加属性尚未支持。");
        foreach (var marker in element.Elements().Where(e => e.Name == Drawing + "from" || e.Name == Drawing + "to"))
            if (marker.HasAttributes || marker.Elements().Count() != 4 || marker.Elements().Select(e => e.Name).Distinct().Count() != 4 ||
                marker.Elements().Any(e => e.Name.Namespace != Drawing || e.Name.LocalName is not ("row" or "col" or "rowOff" or "colOff") || e.HasAttributes || e.HasElements))
                throw new NotSupportedException("绘图标记附加属性尚未支持。");
        long Value(XElement? e, string name, bool attribute = false)
        {
            var text = attribute ? (string?)e?.Attribute(name) : e?.Element(Drawing + name)?.Value;
            if (!long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value < 0 || value > 10_000_000_000_000L)
                throw new InvalidDataException("绘图锚点坐标无效。");
            return value;
        }
        var kind = element.Name.LocalName switch { "twoCellAnchor" => DrawingAnchorKind.TwoCell, "oneCellAnchor" => DrawingAnchorKind.OneCell, "absoluteAnchor" => DrawingAnchorKind.Absolute, _ => throw new NotSupportedException("绘图锚点尚未支持。") };
        var from = element.Element(Drawing + "from"); var to = element.Element(Drawing + "to");
        var extent = element.Element(Drawing + "ext"); var position = element.Element(Drawing + "pos");
        var anchor = new DrawingAnchor { Kind = kind };
        if (kind != DrawingAnchorKind.Absolute) anchor = anchor with { Row = checked((int)Value(from, "row")), Column = checked((int)Value(from, "col")), RowOffset = Value(from, "rowOff"), ColumnOffset = Value(from, "colOff") };
        if (kind == DrawingAnchorKind.TwoCell) anchor = anchor with { EndRow = checked((int)Value(to, "row")), EndColumn = checked((int)Value(to, "col")), EndRowOffset = Value(to, "rowOff"), EndColumnOffset = Value(to, "colOff") };
        else anchor = anchor with { Width = Value(extent, "cx", true), Height = Value(extent, "cy", true) };
        if (kind == DrawingAnchorKind.Absolute) anchor = anchor with { X = Value(position, "x", true), Y = Value(position, "y", true) };
        anchor.Validate();
        // Excel commonly serializes fixed/move-only pictures as two-cell anchors with editAs.
        // Normalize to the equivalent explicit anchor so dimension changes obey the declared behavior.
        if (kind == DrawingAnchorKind.TwoCell && (string?)element.Attribute("editAs") is { } behavior && behavior != "twoCell")
            anchor = DrawingGeometry.FromBounds(geometry, anchor.Bounds(geometry), behavior switch { "oneCell" => DrawingAnchorKind.OneCell, "absolute" => DrawingAnchorKind.Absolute, _ => throw new NotSupportedException("图片跟随单元格规则尚未支持。") });
        if (anchor.Bounds(geometry) is { Width: > 0, Height: > 0 } bounds)
            _ = DrawingGeometry.FromBounds(geometry, bounds, DrawingAnchorKind.TwoCell);
        return anchor;
    }

    internal static XElement Write(DrawingAnchor anchor, XElement content)
    {
        XElement Marker(string name, int row, int col, long rowOff, long colOff) => new(Drawing + name,
            new XElement(Drawing + "col", col), new XElement(Drawing + "colOff", colOff), new XElement(Drawing + "row", row), new XElement(Drawing + "rowOff", rowOff));
        var result = new XElement(Drawing + (anchor.Kind switch { DrawingAnchorKind.TwoCell => "twoCellAnchor", DrawingAnchorKind.OneCell => "oneCellAnchor", _ => "absoluteAnchor" }));
        if (anchor.Kind == DrawingAnchorKind.Absolute) result.Add(new XElement(Drawing + "pos", new XAttribute("x", anchor.X), new XAttribute("y", anchor.Y)));
        else result.Add(Marker("from", anchor.Row, anchor.Column, anchor.RowOffset, anchor.ColumnOffset));
        if (anchor.Kind == DrawingAnchorKind.TwoCell) result.Add(Marker("to", anchor.EndRow, anchor.EndColumn, anchor.EndRowOffset, anchor.EndColumnOffset));
        else result.Add(new XElement(Drawing + "ext", new XAttribute("cx", anchor.Width), new XAttribute("cy", anchor.Height)));
        result.Add(content, new XElement(Drawing + "clientData")); return result;
    }
}
