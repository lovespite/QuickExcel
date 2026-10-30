using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

/// <summary>Only style-local tables and a relationship-free theme may be carried forward.
/// Stable original indexes keep XF, named-style, palette and theme associations intact.</summary>
internal static class XlsxStylePreservation
{
    internal static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly HashSet<string> StyleElements = ("styleSheet numFmts numFmt fonts font name sz color b i u strike outline shadow condense extend family charset scheme vertAlign " +
        "fills fill patternFill gradientFill stop fgColor bgColor borders border left right top bottom diagonal vertical horizontal start end " +
        "cellStyleXfs cellXfs xf alignment protection cellStyles cellStyle colors indexedColors rgbColor mruColors dxfs dxf tableStyles tableStyle tableStyleElement extLst ext").Split(' ').ToHashSet();

    public static bool CanPreserveStyles(XlsxPackage package, string path, XElement root) =>
        package.Relationships(path).Count == 0 && root.DescendantsAndSelf().All(e =>
            e.Attributes().All(a => a.Name.Namespace != Rel) &&
            (e.Name.Namespace == Main && StyleElements.Contains(e.Name.LocalName) && (e.Name.LocalName != "ext" || KnownExtension(e)) || KnownExtensionSetting(e)));

    private static bool KnownExtensionSetting(XElement element) => !element.HasElements &&
        (element.Name == XName.Get("slicerStyles", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main") && element.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == "defaultSlicerStyle") ||
         element.Name == XName.Get("timelineStyles", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main") && element.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == "defaultTimelineStyle"));

    private static bool KnownExtension(XElement element) => element.Elements().Count() == 1 && KnownExtensionSetting(element.Elements().Single()) &&
        (string?)element.Attribute("uri") is "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" or "{9260A510-F301-46a8-8635-F512D64BE5F5}";

    public static bool CanPreserveTheme(XlsxPackage package, string path, XElement root) =>
        package.Relationships(path).Count == 0 && root.Name == Drawing + "theme" && root.DescendantsAndSelf().All(e =>
            e.Attributes().All(a => a.Name.Namespace != Rel) &&
            (e.Name.Namespace == Drawing && (e.Name.LocalName != "ext" || e.Elements().Count() == 1 && KnownThemeFamily(e.Elements().Single()) &&
                (string?)e.Attribute("uri") == "{05A4C25C-085E-4340-85A3-A5531E510DB2}") || KnownThemeFamily(e)));

    private static bool KnownThemeFamily(XElement element) => element.Name == XName.Get("themeFamily", "http://schemas.microsoft.com/office/thememl/2012/main") &&
        !element.HasElements && element.Attributes().All(a => a.IsNamespaceDeclaration || a.Name is { NamespaceName.Length: 0, LocalName: "name" or "id" or "vid" });

    public static (XDocument Styles, byte[]? Theme) Write(WorkbookSnapshot snapshot, CancellationToken cancellationToken)
    {
        var generated = TypedStyles.Write(snapshot.Styles);
        if (snapshot.FileState is not { } state) return (generated, null);
        var source = new XlsxPackage(state.OriginalWorkbook.ToArray(), new(cancellationToken));
        var workbook = source.RelatedByType("", "officeDocument", true)!;
        var stylesPath = source.RelatedByType(workbook, "styles");
        var themePath = source.RelatedByType(workbook, "theme");
        var theme = themePath is not null && CanPreserveTheme(source, themePath, source.XmlPart(themePath).Root!) ? source.BinaryPart(themePath) : null;
        if (stylesPath is null || !CanPreserveStyles(source, stylesPath, source.XmlPart(stylesPath).Root!))
        {
            if (snapshot.Styles.Any(s => s.NamedStyleId != 0)) throw new InvalidOperationException("无法安全保留命名样式关联；请保留原件。");
            return (generated, theme);
        }
        var root = new XElement(source.XmlPart(stylesPath).Root!);
        // Missing parents are already rejected by the reader; empty basic style tables can be regenerated.
        if (root.Element(Main + "cellXfs") is not { } xfs || xfs.Elements().Count() != state.OriginalModel.Styles.Length) return (generated, theme);
        XElement Table(string name)
        {
            if (root.Element(Main + name) is { } table) return table;
            var added = new XElement(Main + name, new XAttribute("count", 0));
            if (name == "numFmts") root.AddFirst(added); else throw new InvalidDataException("样式表缺少 " + name);
            return added;
        }
        var formats = Table("numFmts");
        var usedFormats = formats.Elements().Select(e => TypedStyles.Int(Attr(e, "numFmtId"), 65535)).ToHashSet();
        var originalXfs = xfs.Elements().ToArray();
        xfs.RemoveNodes();
        var freshXfs = generated.Root!.Element(Main + "cellXfs")!.Elements().ToArray();
        for (var i = 0; i < snapshot.Styles.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (i < originalXfs.Length && snapshot.Styles[i] == state.OriginalModel.Styles[i]) { xfs.Add(new XElement(originalXfs[i])); continue; }
            var xf = new XElement(freshXfs[i]);
            foreach (var (tableName, idName) in new[] { ("fonts", "fontId"), ("fills", "fillId"), ("borders", "borderId") })
            {
                var table = Table(tableName);
                var generatedIndex = TypedStyles.Int(Attr(xf, idName), 65535);
                var component = generated.Root.Element(Main + tableName)!.Elements().ElementAt(generatedIndex);
                var index = table.Elements().Count();
                table.Add(new XElement(component)); table.SetAttributeValue("count", index + 1); xf.SetAttributeValue(idName, index);
            }
            if (snapshot.Styles[i].NumberFormatCode is { } code)
            {
                var id = 164; while (usedFormats.Contains(id) && id <= 65535) id++;
                if (id > 65535) throw new InvalidOperationException("自定义数字格式编号已满。");
                usedFormats.Add(id); formats.Add(new XElement(Main + "numFmt", new XAttribute("numFmtId", id), new XAttribute("formatCode", TypedStyles.Encode(code))));
                xf.SetAttributeValue("numFmtId", id);
            }
            xf.SetAttributeValue("xfId", snapshot.Styles[i].NamedStyleId);
            xfs.Add(xf);
        }
        formats.SetAttributeValue("count", formats.Elements().Count());
        xfs.SetAttributeValue("count", snapshot.Styles.Length);
        return (new(root), theme);
    }
}
