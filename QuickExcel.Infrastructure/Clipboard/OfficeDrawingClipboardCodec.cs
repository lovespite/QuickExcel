using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxDrawingAnchors;
using static QuickExcel.Storage.TypedWorkbookXml;

namespace QuickExcel.Clipboard;

/// <summary>Bounded Office GVML chart exchange. Imports cached values, never source links, OLE or executable content.</summary>
public static class OfficeDrawingClipboardCodec
{
    public const int MaxBytes = 16 * 1024 * 1024;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace Locked = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";
    private static FileReadBudget Budget() => new(CancellationToken.None, MaxBytes, 32 * 1024 * 1024);
    private static XlsxPackage Package(byte[] bytes)
    {
        if (bytes.Length is 0 or > MaxBytes) throw new InvalidDataException("对象剪贴板载荷超过 16 MiB 或为空。");
        return new(bytes, Budget(), 128);
    }
    private static string DrawingPath(XlsxPackage package) => package.RelatedByType("", "drawing") ?? throw new NotSupportedException("剪贴板没有 Office 绘图。");
    public static bool ContainsChart(byte[] bytes)
    {
        var package = Package(bytes); var root = package.XmlPart(DrawingPath(package)).Root!;
        return root.Descendants().Any(e => e.Name.LocalName == "chart");
    }
    public static ChartClipboard ReadChart(byte[] bytes)
    {
        var package = Package(bytes); var drawingPath = DrawingPath(package); var drawing = package.XmlPart(drawingPath).Root!;
        ValidateDrawing(drawing);
        var canvas = drawing.Element(Art + "graphicData")?.Element(Locked + "lockedCanvas");
        var frames = canvas?.Elements(Art + "graphicFrame").ToArray() ?? [];
        if (drawing.Name != Art + "graphic" || frames.Length != 1 || canvas!.Elements().Any(e => e.Name != Art + "nvGrpSpPr" && e.Name != Art + "grpSpPr" && e.Name != Art + "graphicFrame"))
            throw new NotSupportedException("可编辑粘贴目前需要单个标准图表；组合对象请选择粘贴为图片。");
        var frame = frames[0]; var transform = frame.Element(Art + "xfrm") ?? throw new InvalidDataException("图表尺寸缺失。");
        if (frame.Elements().Any(e => e.Name != Art + "nvGraphicFramePr" && e.Name != Art + "graphic" && e.Name != Art + "xfrm") ||
            frame.Descendants().Any(e => e.Name.LocalName is "hlinkClick" or "hlinkHover" or "oleObj")) throw new NotSupportedException("图表对象含未支持的附加内容。");
        long Size(XElement t, string name) => long.TryParse((string?)t.Element(Art + "ext")?.Attribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out var v) && v is > 0 and <= 95_250_000
            ? v : throw new InvalidDataException("图表尺寸无效或超过 10,000 像素。");
        var width = Size(transform, "cx"); var height = Size(transform, "cy");
        var rotation = (string?)transform.Attribute("rot") is { } angle && int.TryParse(angle, out var degrees) && degrees is >= 0 and < 21_600_000 ? degrees / 60000d : transform.Attribute("rot") is null ? 0 : throw new InvalidDataException("图表旋转角度无效。");
        foreach (var t in new[] { transform, canvas.Element(Art + "grpSpPr")?.Element(Art + "xfrm") }.OfType<XElement>())
        {
            if (t.Attributes().Any(a => !a.IsNamespaceDeclaration && !(t == transform && a.Name == "rot") && a.Value is not ("0" or "false")))
                throw new NotSupportedException("翻转或旋转组合图表请选择粘贴为图片。");
            if (t != transform && (Size(t, "cx") != width || Size(t, "cy") != height ||
                (string?)t.Element(Art + "chExt")?.Attribute("cx") != width.ToString(CultureInfo.InvariantCulture) ||
                (string?)t.Element(Art + "chExt")?.Attribute("cy") != height.ToString(CultureInfo.InvariantCulture))) throw new NotSupportedException("缩放组合图表请选择粘贴为图片。");
        }
        var reference = frame.Element(Art + "graphic")?.Element(Art + "graphicData")?.Element(C + "chart") ?? throw new NotSupportedException("此图表类型尚不支持可编辑粘贴；请选择粘贴为图片。");
        var chartPath = package.Related(drawingPath, XlsxPackage.Attr(reference, Rel + "id"), "chart");
        var xml = new XElement(package.XmlPart(chartPath).Root!);
        if (xml.Descendants().Any(e => e.Name.NamespaceName == "http://schemas.microsoft.com/office/drawing/2012/chart"))
            throw new NotSupportedException("外部筛选图表的隐藏数据不在剪贴板缓存中；请选择粘贴为图片，或在源应用取消图表筛选后复制。");
        if (package.Parts[chartPath].Length > 512 * 1024) throw new InvalidDataException("图表定义超过 512 KiB。");
        var dateSystem = (string?)xml.Element(C + "date1904")?.Attribute("val") switch { null or "0" or "false" => ExcelDateSystem.Windows1900, "1" or "true" => ExcelDateSystem.Mac1904, _ => throw new InvalidDataException("图表日期系统无效。") };
        var data = new WorksheetSnapshot(SheetId.New(), "图表数据"); var cells = new Dictionary<CellPosition, CellSnapshot>();
        var styles = new List<StyleDefinition> { new() }; var mapped = new Dictionary<string, (string Formula, string Cache)>();
        var literalIndex = 0;
        foreach (var literal in xml.Descendants().Where(e => e.Name == C + "numLit" || e.Name == C + "strLit").ToArray())
        {
            var numeric = literal.Name == C + "numLit";
            var cache = new XElement(literal) { Name = C + (numeric ? "numCache" : "strCache") };
            literal.ReplaceWith(new XElement(C + (numeric ? "numRef" : "strRef"), new XElement(C + "f", "clipboard-literal-" + literalIndex++), cache));
        }
        foreach (var source in xml.Descendants().Where(e => e.Name == C + "strRef" || e.Name == C + "numRef").ToArray())
        {
            var numeric = source.Name == C + "numRef"; var cache = source.Element(C + (numeric ? "numCache" : "strCache"));
            var original = source.Element(C + "f") ?? throw new InvalidDataException("图表源引用缺失。");
            if (original.Value.Length > 4096 || cache is null || !int.TryParse((string?)cache.Element(C + "ptCount")?.Attribute("val"), out var count) || count is < 1 or > 4096)
                throw new NotSupportedException("图表数据缓存缺失或超限；请选择粘贴为图片。");
            var cacheText = cache.ToString(SaveOptions.DisableFormatting); var key = source.Name.LocalName + ":" + original.Value;
            if (mapped.TryGetValue(key, out var prior))
            {
                if (prior.Cache != cacheText) throw new InvalidDataException("同一图表引用包含不一致的数据缓存。");
                original.Value = prior.Formula; continue;
            }
            if (cells.Count + count > ChartClipboard.MaxCells) throw new NotSupportedException("图表数据超过 20,000 格。");
            var column = mapped.Count; var values = new CellValue[count]; var indices = new HashSet<int>();
            foreach (var point in cache.Elements(C + "pt"))
            {
                if (!int.TryParse((string?)point.Attribute("idx"), out var index) || index < 0 || index >= count || !indices.Add(index) || point.Element(C + "v") is not { } value)
                    throw new InvalidDataException("图表数据缓存索引无效。");
                if (numeric)
                {
                    if (!double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)) throw new NotSupportedException("图表含非数值缓存；请选择粘贴为图片。");
                    values[index] = CellValue.FromNumber(number);
                }
                else values[index] = CellValue.FromText(value.Value);
            }
            var style = new StyleDefinition { NumberFormatCode = numeric ? cache.Element(C + "formatCode")?.Value ?? "General" : "General" };
            var styleId = styles.IndexOf(style); if (styleId < 0) { styleId = styles.Count; styles.Add(style); }
            for (var i = 0; i < count; i++) cells.Add(new(i, column), new(CellContent.FromLiteral(values[i], styleId)));
            original.Value = FormulaSyntax.Print(new ReferenceNode(data.Name, 0, column, count - 1, column, true, true, true, true, IsRange: count > 1)).TrimStart('=');
            mapped.Add(key, (original.Value, cacheText));
        }
        // Cached values are intentionally detached from external files/embedded workbooks. Never open those relationships.
        xml.Elements(C + "externalData").Remove();
        var parts = new Dictionary<string, byte[]>(package.Parts) { [chartPath] = XmlBytes(new(xml)) };
        var chartRels = package.Relationships(chartPath);
        if (chartRels.Values.Any(r => !r.Type.EndsWith("/oleObject", StringComparison.Ordinal) && !r.Type.EndsWith("/package", StringComparison.Ordinal) && !r.Type.EndsWith("/chartStyle", StringComparison.Ordinal) && !r.Type.EndsWith("/chartColorStyle", StringComparison.Ordinal)))
            throw new NotSupportedException("图表含未支持的附加关系；请选择粘贴为图片。");
        if (parts.TryGetValue(RelationshipPart(chartPath), out var rawRels))
        {
            var rels = Xml(rawRels, Budget()); rels.Root!.Elements().Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/oleObject", StringComparison.Ordinal) == true || ((string?)r.Attribute("Type"))?.EndsWith("/package", StringComparison.Ordinal) == true).Remove();
            parts[RelationshipPart(chartPath)] = XmlBytes(rels);
        }
        if (package.RelatedByType(drawingPath, "theme") is { } themePath)
        {
            var theme = new XElement(package.XmlPart(themePath).Root!);
            if (theme.Name == Art + "clipboardTheme")
            {
                var map = theme.Element(Art + "clrMap");
                if (map?.HasElements == true || map?.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Value != (a.Name.LocalName switch { "bg1" => "lt1", "tx1" => "dk1", "bg2" => "lt2", "tx2" => "dk2", _ => a.Name.LocalName })) == true)
                    throw new NotSupportedException("自定义主题颜色映射请选择粘贴为图片。");
                theme.Name = Art + "theme"; theme.SetAttributeValue("name", "Clipboard"); theme.Elements(Art + "clrMap").Remove();
            }
            parts[themePath] = XmlBytes(new(theme));
        }
        var anchor = new DrawingAnchor { Kind = DrawingAnchorKind.OneCell, Width = width, Height = height, Rotation = rotation };
        var converted = new XElement(Drawing + "graphicFrame", new XElement(Drawing + "nvGraphicFramePr", new XElement(Drawing + "cNvPr", new XAttribute("id", 1), new XAttribute("name", "Chart")), new XElement(Drawing + "cNvGraphicFramePr")),
            new XElement(Drawing + "xfrm", new XElement(Art + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(Art + "ext", new XAttribute("cx", width), new XAttribute("cy", height))), new XElement(frame.Element(Art + "graphic")!));
        parts[drawingPath] = XmlBytes(new(new XElement(Drawing + "wsDr", XlsxDrawingAnchors.Write(anchor, converted))));
        const string sheetPath = "quickexcel/clipboard.xml";
        var worksheet = new XElement(Main + "worksheet", new XElement(Main + "drawing", new XAttribute(Rel + "id", "drawing")));
        parts[sheetPath] = XmlBytes(new(worksheet)); parts[RelationshipPart(sheetPath)] = XmlBytes(new(new XElement(PackageRel + "Relationships", Relationship("drawing", "drawing", "../" + drawingPath))));
        var parsed = new XlsxPackage(Zip(parts), Budget(), 132);
        data = new(data.Id, data.Name, cells);
        var charts = XlsxCharts.Read(parsed, drawingPath, sheetPath, worksheet, data);
        if (charts.Length != 1 || parsed.Issues.Count > 0) throw new NotSupportedException("此图表无法完整转换为可编辑图表；请选择粘贴为图片。" + string.Join("；", parsed.Issues));
        data = data.WithData(data.Data with { Charts = charts });
        return ChartClipboard.Capture(new(Guid.NewGuid(), 0, dateSystem, [data], styles: styles), charts[0]);
    }

    private static void ValidateDrawing(XElement drawing)
    {
        foreach (var e in drawing.DescendantsAndSelf())
        {
            var allowed = e.Name.LocalName switch
            {
                "graphic" or "lockedCanvas" or "nvGrpSpPr" or "cNvGrpSpPr" or "grpSpPr" or "graphicFrame" or "nvGraphicFramePr" or "cNvGraphicFramePr" or "graphicFrameLocks" or "extLst" => "",
                "graphicData" => "uri", "cNvPr" => "id name", "xfrm" => "rot flipH flipV", "off" or "chOff" => "x y", "chExt" => "cx cy",
                "ext" => e.Parent?.Name == Art + "extLst" ? "uri" : "cx cy", "chart" => "id", "creationId" => "id", _ => null
            };
            XNamespace identity = "http://schemas.microsoft.com/office/drawing/2014/main";
            if (allowed is null || e.Name.Namespace != Art && e.Name != Locked + "lockedCanvas" && e.Name != C + "chart" && e.Name != identity + "creationId" ||
                e.Attributes().Any(a => !a.IsNamespaceDeclaration && (!allowed.Split(' ').Contains(a.Name.LocalName) || a.Name.Namespace != XNamespace.None && !(e.Name == C + "chart" && a.Name == Rel + "id"))))
                throw new NotSupportedException("绘图含未支持的对象属性；请选择粘贴为图片。");
            if (e.Parent?.Name == Art + "extLst" && ((string?)e.Attribute("uri") != "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" || e.Elements().Count() != 1 || e.Element(identity + "creationId") is not { } id || !Guid.TryParse((string?)id.Attribute("id"), out _)))
                throw new NotSupportedException("绘图含未知扩展；请选择粘贴为图片。");
        }
    }

    public static byte[] WritePrivateChart(ChartClipboard source)
    {
        var book = source.Book.WithSheets([source.Book.Sheets[0].WithData(source.Book.Sheets[0].Data with { Charts = [source.Chart] })]);
        var bytes = TypedWorkbookXml.Write(book, false, CancellationToken.None);
        _ = Package(bytes); // Enforce the same expanded-size and entry budgets as the receiver.
        return bytes;
    }
    /// <summary>Office charts use literal series arrays, so pasting cannot bind to unrelated destination cells.</summary>
    public static byte[]? WriteOfficeChart(ChartClipboard source)
    {
        if (source.Chart.IsExtended || source.Chart.Series.Any(s => s.IsFiltered || !s.HiddenPoints.IsEmpty)) return null;
        var parts = Unzip(WritePrivateChart(source), Budget(), 128);
        const string drawingPath = "xl/drawings/drawing1.xml", chartPath = "xl/charts/chart1-1.xml";
        var chartXml = Xml(parts[chartPath], Budget());
        foreach (var reference in chartXml.Descendants().Where(e => e.Name == C + "numRef" || e.Name == C + "strRef").ToArray())
        {
            var numeric = reference.Name == C + "numRef"; var cache = reference.Element(C + (numeric ? "numCache" : "strCache"))!;
            if (reference.Parent!.Name == C + "tx") reference.ReplaceWith(new XElement(C + "v", cache.Element(C + "pt")?.Element(C + "v")?.Value ?? ""));
            else reference.ReplaceWith(new XElement(C + (numeric ? "numLit" : "strLit"), cache.Elements()));
        }
        parts[chartPath] = XmlBytes(chartXml);
        var old = Xml(parts[drawingPath], Budget()).Root!.Elements().Single().Element(Drawing + "graphicFrame")!;
        var b = source.Chart.Anchor.Bounds(source.Book.Sheets[0]); var width = (long)Math.Round(b.Width * DrawingAnchor.EmusPerPixel); var height = (long)Math.Round(b.Height * DrawingAnchor.EmusPerPixel);
        XElement Offset(string name) => new(Art + name, new XAttribute("x", 0), new XAttribute("y", 0));
        XElement Extent(string name) => new(Art + name, new XAttribute("cx", width), new XAttribute("cy", height));
        // GVML uses nvGraphicFramePr, graphic, xfrm (a different order from spreadsheetDrawing).
        var frame = new XElement(Art + "graphicFrame", new XElement(Art + "nvGraphicFramePr", new XElement(Art + "cNvPr", new XAttribute("id", 1), new XAttribute("name", "Chart")), new XElement(Art + "cNvGraphicFramePr")),
            new XElement(old.Element(Art + "graphic")!), new XElement(Art + "xfrm", new XAttribute("rot", (int)Math.Round(source.Chart.Anchor.Rotation * 60000)), Offset("off"), Extent("ext")));
        parts[drawingPath] = XmlBytes(new(new XElement(Art + "graphic", new XElement(Art + "graphicData", new XAttribute("uri", Locked.NamespaceName), new XElement(Locked + "lockedCanvas",
            new XElement(Art + "nvGrpSpPr", new XElement(Art + "cNvPr", new XAttribute("id", 0), new XAttribute("name", "")), new XElement(Art + "cNvGrpSpPr")),
            new XElement(Art + "grpSpPr", new XElement(Art + "xfrm", Offset("off"), Extent("ext"), Offset("chOff"), Extent("chExt"))), frame)))));
        parts["_rels/.rels"] = XmlBytes(new(new XElement(PackageRel + "Relationships", Relationship("drawing", "drawing", drawingPath))));
        var bytes = Zip(parts); if (bytes.Length > MaxBytes) throw new NotSupportedException("图表剪贴板超过容量限制。"); return bytes;
    }
    public static ChartClipboard ReadPrivateChart(byte[] bytes)
    {
        _ = Package(bytes);
        var parsed = TypedWorkbookXml.Read(bytes, Budget()); var book = parsed.Snapshot;
        if (parsed.Issues.Count > 0 || book.Sheets.Length != 1 || book.Sheets[0].Data.Charts.Length != 1 || book.Sheets[0].Cells.Count > ChartClipboard.MaxCells ||
            book.Sheets[0].Cells.Values.Any(c => c.Content.Kind != CellContentKind.Literal) || book.Sheets[0].Images.Length > 0)
            throw new NotSupportedException("图表剪贴板必须只包含一个可建模图表及其结果值，不能包含公式或未知内容。");
        return ChartClipboard.Capture(book, book.Sheets[0].Data.Charts[0]);
    }
}
