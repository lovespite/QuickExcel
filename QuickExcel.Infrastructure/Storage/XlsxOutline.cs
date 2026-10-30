using System.Xml.Linq;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

internal static class XlsxOutline
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal static OutlineSettings? Read(XlsxPackage package, string path, XElement sheet)
    {
        var element = sheet.Element(Main + "sheetPr")?.Element(Main + "outlinePr");
        if (element is null) return null;
        package.Shape(element, "", "summaryBelow summaryRight showOutlineSymbols applyStyles", path);
        bool Flag(string name, bool fallback) => element.Attribute(name) is not { } a ? fallback : a.Value is "1" or "true";
        return new(Flag("summaryBelow", true), Flag("summaryRight", true), Flag("showOutlineSymbols", true), Flag("applyStyles", false));
    }
    internal static XElement? Write(OutlineSettings? settings) => settings is null ? null : new(Main + "outlinePr",
        new XAttribute("summaryBelow", settings.SummaryBelow ? 1 : 0), new XAttribute("summaryRight", settings.SummaryRight ? 1 : 0),
        new XAttribute("showOutlineSymbols", settings.ShowSymbols ? 1 : 0), new XAttribute("applyStyles", settings.ApplyStyles ? 1 : 0));
}
