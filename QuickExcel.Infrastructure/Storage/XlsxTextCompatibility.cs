using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.XlsxPackage;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

internal static class XlsxTextCompatibility
{
    private const string Extension = "{D14903EA-33C4-47F7-8F05-3474C54BE107}";
    private static readonly XNamespace Version = "http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion";
    public static int Read(XlsxPackage package, XElement root, string part)
    {
        var extensions = root.Elements(Main + "extLst").Elements(Main + "ext").Where(e => (string?)e.Attribute("uri") == Extension).ToArray();
        if (extensions.Length == 0) return 1;
        if (extensions.Length != 1 || extensions[0].Elements().Count() != 1 || extensions[0].Elements().Single().Name != Version + "version")
            throw new InvalidDataException("文本计算兼容版本扩展无效。");
        var element = extensions[0].Elements().Single();
        if (element.HasElements || !string.IsNullOrWhiteSpace(element.Value) || element.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "setVersion" && a.Name != "warnBelowVersion"))
            throw new InvalidDataException("文本计算兼容版本字段无效。");
        if (!int.TryParse((string?)element.Attribute("setVersion"), NumberStyles.None, CultureInfo.InvariantCulture, out var version) || version < 1)
            throw new InvalidDataException("文本计算兼容版本必须为正整数。");
        if (element.Attribute("warnBelowVersion") is { } warning && (!int.TryParse(warning.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var below) || below < 1 || below > version))
            throw new InvalidDataException("文本计算兼容警告版本无效。");
        if (version >= 2 && (string?)element.Attribute("warnBelowVersion") != version.ToString(CultureInfo.InvariantCulture))
            package.Loss("兼容警告阈值未建模；受支持副本以当前版本为阈值。", "workbookCompatibilityVersion/warning", part);
        if (extensions[0].Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "uri"))
            package.Loss("文本兼容扩展有未建模属性。", "workbookCompatibilityVersion/extension", part);
        if (version > 2) package.Loss("未知文本计算兼容版本；公式只读保留。", "workbookCompatibilityVersion", part);
        return version;
    }
    public static XElement Shape(XElement root)
    {
        var result = new XElement(root);
        result.Elements(Main + "extLst").Elements(Main + "ext").Where(e => (string?)e.Attribute("uri") == Extension).Remove();
        result.Elements(Main + "extLst").Where(e => !e.HasElements && !e.HasAttributes).Remove();
        return result;
    }
    public static XElement? Write(WorkbookSnapshot snapshot) => snapshot.TextCompatibilityVersion == 1 ? null : new(Main + "extLst",
        new XElement(Main + "ext", new XAttribute("uri", Extension),
            new XElement(Version + "version", new XAttribute(XNamespace.Xmlns + "xlwcv", Version),
                new XAttribute("setVersion", snapshot.TextCompatibilityVersion), new XAttribute("warnBelowVersion", snapshot.TextCompatibilityVersion))));
}
