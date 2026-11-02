using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.XlsxPackage;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

internal static class XlsxProtection
{
    private static readonly System.Reflection.PropertyInfo[] Flags = typeof(WorksheetProtection).GetProperties().Where(p => p.PropertyType == typeof(bool) && p.CanWrite).ToArray();
    private static string Name(string name) => char.ToLowerInvariant(name[0]) + name[1..];
    internal static WorksheetProtection? Read(XlsxPackage package, string path, XElement sheet)
    {
        if (sheet.Elements(Main + "sheetProtection").Count() > 1) throw new InvalidDataException("重复的工作表保护定义。");
        if (sheet.Element(Main + "sheetProtection") is not { } element) return null;
        package.Shape(element, "", string.Join(" ", Flags.Select(p => Name(p.Name))) + " password algorithmName hashValue saltValue spinCount", path);
        var result = new WorksheetProtection { Sheet = false, Objects = false, Scenarios = false,
            Password = (string?)element.Attribute("password"), AlgorithmName = (string?)element.Attribute("algorithmName"),
            HashValue = (string?)element.Attribute("hashValue"), SaltValue = (string?)element.Attribute("saltValue"), SpinCount = (uint?)element.Attribute("spinCount") };
        foreach (var flag in Flags) if (element.Attribute(Name(flag.Name)) is { } attribute) flag.SetValue(result, TypedStyles.Bool(attribute.Value));
        return result;
    }
    internal static XElement? Write(WorksheetProtection? protection)
    {
        if (protection is null) return null;
        var result = new XElement(Main + "sheetProtection", Flags.Select(p => new XAttribute(Name(p.Name), (bool)p.GetValue(protection)! ? 1 : 0)));
        foreach (var p in typeof(WorksheetProtection).GetProperties().Where(p => p.CanWrite && p.PropertyType != typeof(bool)))
            if (p.GetValue(protection) is { } value) result.Add(new XAttribute(Name(p.Name), value));
        return result;
    }
}
