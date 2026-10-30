using System.Text;
using System.Xml.Linq;
using static QuickExcel.CompatibilityProbe.OpenXmlAudit;

namespace QuickExcel.CompatibilityProbe;

public sealed record CompatibilityFixture(string Id, string Scenario, string[] RoadmapItems, byte[] Bytes);

/// <summary>Synthetic, independently authored SpreadsheetML. These are not Excel-authored acceptance evidence.</summary>
public static class FixtureCatalog
{
    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);
    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """;

    public static IEnumerable<CompatibilityFixture> Create()
    {
        yield return Fixture("roster", "名单：文本编号、空文本、布尔、中文", ["XC-0107"], Base("""
            <row r="1"><c r="A1" t="inlineStr"><is><t>编号</t></is></c><c r="B1" t="inlineStr"><is><t>姓名</t></is></c></row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>00123</t></is></c><c r="B2" t="inlineStr"><is><t>张三</t></is></c><c r="C2" t="b"><v>1</v></c><c r="D2" t="inlineStr"><is><t></t></is></c></row>
            """));

        var quote = Base("<row r=\"2\"><c r=\"A2\" s=\"2\"><v>1234.567</v></c><c r=\"B2\" s=\"2\"><v>-42.5</v></c></row>");
        Edit(quote, "xl/styles.xml", root =>
        {
            root.AddFirst(new XElement(Main + "numFmts", new XAttribute("count", 1),
                new XElement(Main + "numFmt", new XAttribute("numFmtId", 164), new XAttribute("formatCode", "¥#,##0.000;[Red]-¥#,##0.000;\"—\";@"))));
            root.Element(Main + "fonts")!.Elements().Single().Add(new XElement(Main + "scheme", new XAttribute("val", "minor")));
            var xfs = root.Element(Main + "cellXfs")!;
            xfs.SetAttributeValue("count", 3);
            xfs.Add(new XElement(Main + "xf", new XAttribute("numFmtId", 164), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0),
                new XAttribute("xfId", 0), new XAttribute("applyFont", 0), new XAttribute("applyNumberFormat", 1)));
        });
        yield return Fixture("quotation", "报价：货币分段、小数、主题字体及 apply 标志", ["XC-0101", "XC-0102"], quote);

        foreach (var mac in new[] { false, true })
        {
            var dates = Base("<row r=\"2\"><c r=\"A2\" s=\"1\"><v>60</v></c><c r=\"B2\" s=\"1\"><v>0</v></c><c r=\"C2\" s=\"1\"><v>45292.5</v></c></row>");
            Edit(dates, "xl/workbook.xml", root => root.Element(Main + "workbookPr")!.SetAttributeValue("date1904", mac ? "1" : "0"));
            yield return Fixture(mac ? "date-ledger-1904" : "date-ledger-1900", "日期台账：原始序列、60、0、小数", ["XC-0101"], dates);
        }

        var summary = Base("<row r=\"2\"><c r=\"A2\"><v>21</v></c></row>", "<row r=\"2\"><c r=\"A2\"><f>'数据'!A2*2</f><v>42</v></c></row>");
        yield return Fixture("cross-sheet", "跨表汇总：公式与旧缓存分离", ["XC-0106"], summary);

        var print = Base("<row r=\"2\"><c r=\"A2\"><v>10</v></c></row>", "");
        Edit(print, "xl/workbook.xml", root =>
        {
            root.Element(Main + "sheets")!.AddBeforeSelf(new XElement(Main + "bookViews", new XElement(Main + "workbookView", new XAttribute("activeTab", 0))));
            root.Element(Main + "sheets")!.Elements().Last().SetAttributeValue("state", "veryHidden");
            root.Add(new XElement(Main + "definedNames",
                new XElement(Main + "definedName", new XAttribute("name", "_xlnm.Print_Area"), new XAttribute("localSheetId", 0), "'数据'!$A$1:$D$20"),
                new XElement(Main + "definedName", new XAttribute("name", "_xlnm.Print_Titles"), new XAttribute("localSheetId", 0), "'数据'!$1:$1")));
        });
        Edit(print, "xl/worksheets/sheet1.xml", root =>
        {
            root.AddFirst(new XElement(Main + "sheetViews", new XElement(Main + "sheetView", new XAttribute("workbookViewId", 0), new XAttribute("zoomScale", 125),
                new XElement(Main + "pane", new XAttribute("ySplit", 1), new XAttribute("topLeftCell", "A2"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen")),
                new XElement(Main + "selection", new XAttribute("pane", "bottomLeft"), new XAttribute("activeCell", "B2"), new XAttribute("sqref", "B2")))));
            root.Add(new XElement(Main + "pageMargins", new XAttribute("left", 0.7), new XAttribute("right", 0.7), new XAttribute("top", 0.75), new XAttribute("bottom", 0.75), new XAttribute("header", 0.3), new XAttribute("footer", 0.3)),
                new XElement(Main + "pageSetup", new XAttribute("paperSize", 9), new XAttribute("orientation", "landscape"), new XAttribute("fitToWidth", 1)),
                new XElement(Main + "headerFooter", new XElement(Main + "oddHeader", "&C报价单"), new XElement(Main + "oddFooter", "&P / &N")));
        });
        AddPart(print, "docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", """
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>合成打印模板</dc:title><dc:creator>QuickExcel compatibility fixture</dc:creator></cp:coreProperties>
            """);
        Edit(print, "_rels/.rels", root => root.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "props"),
            new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"), new XAttribute("Target", "docProps/core.xml"))));
        yield return Fixture("print-template", "打印模板：属性、冻结/缩放/活动格、veryHidden、打印定义", ["XC-0103", "XC-0104", "XC-0105"], print);

        var protectedPrint = print.ToDictionary(p => p.Key, p => p.Value);
        Edit(protectedPrint, "xl/worksheets/sheet1.xml", root => root.Element(Main + "sheetData")!.AddAfterSelf(new XElement(Main + "sheetProtection", new XAttribute("sheet", 1))));
        yield return Fixture("protected-print-template", "受保护打印模板：页面设置已支持，工作表保护仍需报告", ["XC-0404"], protectedPrint);

        yield return Fixture("unknown-function", "未知函数与版本前缀：源码和旧缓存不等于可计算", ["XC-0106"], Base("""
            <row r="2"><c r="A2"><f>_xlfn.XLOOKUP(1,B2:B3,C2:C3)</f><v>7</v></c><c r="B2"><v>1</v></c><c r="C2"><v>7</v></c></row>
            """));
        yield return Fixture("array-formula", "数组公式：只读查看来源缓存、整包保留，拒绝编辑和公式删减副本", ["XC-0106"], Base("""
            <row r="2"><c r="B2"><f t="array" ref="B2:B3">ROW(A1:A2)</f><v>1</v></c></row>
            <row r="3"><c r="B3"><v>2</v></c></row>
            """));

        var rich = Base("""
            <row r="2"><c r="A2" t="inlineStr"><is><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>红色</t></r><r><t>文字</t></r></is></c></row>
            """);
        Edit(rich, "xl/worksheets/sheet1.xml", root => root.Add(new XElement(Main + "hyperlinks",
            new XElement(Main + "hyperlink", new XAttribute("ref", "A2"), new XAttribute(Rel + "id", "link")))));
        rich["xl/worksheets/_rels/sheet1.xml.rels"] = Utf8($"""
            <Relationships xmlns="{PackageRel}"><Relationship Id="link" Type="{Rel}/hyperlink" Target="https://example.invalid/fixture" TargetMode="External"/><Relationship Id="notes" Type="{Rel}/comments" Target="../comments1.xml"/></Relationships>
            """);
        AddPart(rich, "xl/comments1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", $"""
            <comments xmlns="{Main}"><authors><author>Fixture</author></authors><commentList><comment ref="A2" authorId="0"><text><t>传统批注</t></text></comment></commentList></comments>
            """);
        yield return Fixture("rich-links-notes", "富文本、外部超链接、传统批注", ["XC-0107"], rich);

        var rowStyles = Base("""
            <row r="1" ht="24" customHeight="1"><c r="A1"><v>61</v></c><c r="B1" s="0"><v>61</v></c></row>
            <row r="2" s="0" customFormat="1"><c r="A2"><v>61</v></c><c r="B2" s="2"><v>0.5</v></c></row>
            <row r="3" s="2" customFormat="0"><c r="A3"><v>61</v></c></row>
            <row r="4" s="2"><c r="A4"><v>61</v></c></row>
            <row r="5" s="2" customFormat="true"><c r="A5"><v>0.5</v></c><c r="B5" s="0"><v>61</v></c></row>
            <row r="6" customFormat="true"><c r="A6"><v>61</v></c></row>
            <row r="7" s="0" customFormat="false"/>
            """);
        Edit(rowStyles, "xl/styles.xml", root =>
        {
            var xfs = root.Element(Main + "cellXfs")!;
            xfs.SetAttributeValue("count", 3);
            var time = new XElement(xfs.Elements().Last()); time.SetAttributeValue("numFmtId", 21); xfs.Add(time);
        });
        Edit(rowStyles, "xl/worksheets/sheet1.xml", root => root.AddFirst(new XElement(Main + "cols",
            new XElement(Main + "col", new XAttribute("min", 1), new XAttribute("max", 4), new XAttribute("style", 1)))));
        yield return Fixture("row-column-styles", "行列继承：显式零样式、禁用/缺省行格式、尺寸行与单格覆盖", ["XC-0102"], rowStyles);
    }

    private static CompatibilityFixture Fixture(string id, string scenario, string[] items, Dictionary<string, byte[]> parts) => new(id, scenario, items, Zip(parts));

    private static Dictionary<string, byte[]> Base(string rows, string? secondRows = null)
    {
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["[Content_Types].xml"] = Utf8("""
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>
                """),
            ["_rels/.rels"] = Utf8($"<Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"workbook\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"),
            ["xl/_rels/workbook.xml.rels"] = Utf8($"<Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"styles\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/></Relationships>")
        };
        AddPart(parts, "xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><workbookPr date1904=\"0\"/><sheets/></workbook>");
        AddPart(parts, "xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", Styles);
        var sheets = secondRows is null ? new[] { rows } : new[] { rows, secondRows };
        for (var i = 0; i < sheets.Length; i++)
        {
            var id = i + 1;
            AddPart(parts, $"xl/worksheets/sheet{id}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                $"<worksheet xmlns=\"{Main}\"><sheetData>{sheets[i]}</sheetData></worksheet>");
            Edit(parts, "xl/workbook.xml", root => root.Element(Main + "sheets")!.Add(new XElement(Main + "sheet",
                new XAttribute("name", i == 0 ? "数据" : "汇总"), new XAttribute("sheetId", id), new XAttribute(Rel + "id", $"sheet{id}"))));
            Edit(parts, "xl/_rels/workbook.xml.rels", root => root.Add(new XElement(PackageRel + "Relationship",
                new XAttribute("Id", $"sheet{id}"), new XAttribute("Type", Rel.NamespaceName + "/worksheet"), new XAttribute("Target", $"worksheets/sheet{id}.xml"))));
        }
        return parts;
    }

    private static void AddPart(Dictionary<string, byte[]> parts, string path, string type, string content)
    {
        parts.Add(path, Utf8(content));
        Edit(parts, "[Content_Types].xml", root => root.Add(new XElement(root.Name.Namespace + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", type))));
    }

    public static void Edit(Dictionary<string, byte[]> parts, string part, Action<XElement> edit)
    {
        var document = Xml(parts[part]);
        edit(document.Root!);
        parts[part] = Utf8(document.ToString(SaveOptions.DisableFormatting));
    }
}
