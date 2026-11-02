using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace QuickExcel.Diagnostics;

/// <summary>Independent ZIP/XML fixtures for desktop smoke and linked storage tests; not a workbook serializer.</summary>
internal static class WorkbookPackageFixture
{
    internal static byte[] Xlsx(string sheetData, IReadOnlyDictionary<string, string>? overrides = null)
    {
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """,
            ["xl/workbook.xml"] = """
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """,
            ["xl/_rels/workbook.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """,
            ["xl/worksheets/sheet1.xml"] = $"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>{sheetData}</sheetData></worksheet>"
        };
        if (overrides is not null)
            foreach (var (name, xml) in overrides) parts[name] = xml;
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var (name, xml) in parts) Write(zip, name, Encoding.UTF8.GetBytes(xml));
        return output.ToArray();
    }

    internal static void WriteV1(string path, byte[] workbook, string script = "")
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        Write(zip, "manifest.json", Encoding.UTF8.GetBytes("""
            {"format":"QuickExcel","formatVersion":1,"scriptApiVersion":1,"workbook":"workbook.xlsx","script":"scripts/functions.js"}
            """));
        Write(zip, "workbook.xlsx", workbook);
        Write(zip, "scripts/functions.js", Encoding.UTF8.GetBytes(script));
    }

    private static void Write(ZipArchive zip, string name, byte[] bytes)
    {
        using var stream = zip.CreateEntry(name).Open();
        stream.Write(bytes);
    }
}
