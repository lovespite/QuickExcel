using System.Text;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed partial class TypedStorageTests
{
    [Theory]
    [InlineData(30.4)]
    [InlineData(30.1)]
    [InlineData(0.1)]
    public void FractionalModelHeightsSurviveSaveAndExportVerification(double height)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Fractional", rows: [new(0, new(height))], defaultRowHeight: height);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        var path = PathFor("fractional.xlsq"); _storage.Save(path, snapshot);
        Assert.Equal(height, _storage.Open(path).Sheets[0].Rows[0].Height!.Value, 10);
        _storage.Export(PathFor("fractional.xlsx"), snapshot);
        Assert.Equal(height, _storage.Open(PathFor("fractional.xlsx")).Sheets[0].DefaultRowHeight!.Value, 10);
    }

    [Fact]
    public void XlsxRowPointsConvertToModelPixelsAndExportBackToPoints()
    {
        var path = PathFor("row-points.xlsx"); _storage.Export(path, Model([Literal(CellValue.FromText("18 point text"))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/worksheets/sheet1.xml"]));
        xml.Root!.Add(new XElement(Ns + "sheetFormatPr", new XAttribute("defaultRowHeight", "22.8")));
        xml.Descendants(Ns + "row").Single().SetAttributeValue("ht", "25.2");
        parts["xl/worksheets/sheet1.xml"] = Encoding.UTF8.GetBytes(xml.ToString()); File.WriteAllBytes(path, Zip(parts));
        var snapshot = _storage.Open(path);
        Assert.Equal(30.4, snapshot.Sheets[0].DefaultRowHeight!.Value, 10);
        Assert.Equal(33.6, snapshot.Sheets[0].Rows[0].Height!.Value, 10);
        var output = PathFor("points-roundtrip.xlsx");
        _storage.Export(output, snapshot);
        var written = XDocument.Parse(Encoding.UTF8.GetString(ReadZip(File.ReadAllBytes(output))["xl/worksheets/sheet1.xml"]));
        Assert.Equal(22.8, (double)written.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultRowHeight")!, 10);
        Assert.Equal(25.2, (double)written.Descendants(Ns + "row").Single().Attribute("ht")!, 10);
        var package = PathFor("dimensions.xlsq");
        _storage.Save(package, snapshot, true);
        Assert.Equal(snapshot.Sheets[0].Rows, _storage.Open(package).Sheets[0].Rows);
    }

    [Theory]
    [InlineData("theme=0", "FFFFFFFF")]
    [InlineData("theme=1", "FF000000")]
    [InlineData("theme=2", "FFE7E6E6")]
    [InlineData("theme=3", "FF44546A")]
    [InlineData("theme=5", "FFED7D31")]
    [InlineData("theme=9;tint=-0.249977111117893", "FF548235")]
    [InlineData("theme=0;tint=-4.9989318521683403E-2", "FFF2F2F2")]
    [InlineData("theme=4;tint=1", "FFFFFFFF")]
    [InlineData("theme=4;tint=-1", "FF000000")]
    [InlineData("rgb=FF336699;tint=0.5", "FF8CB2D9")]
    [InlineData("indexed=10", "FFFF0000")]
    [InlineData("indexed=64", "FF000000")]
    [InlineData("indexed=65", "FFFFFFFF")]
    [InlineData("auto=1", "FF000000")]
    public void ThemeIndexedAndTintedColorsAreResolvedForTextFillAndBorder(string attributes, string expected)
    {
        var path = ColorFixture(attributes);
        var snapshot = _storage.Open(path); var style = snapshot.Styles[0];
        Assert.Equal(expected, style.FontColor);
        Assert.Equal(expected, style.FillColor);
        Assert.Equal(expected, style.BottomBorderColor);
        Assert.DoesNotContain(snapshot.FileState!.CompatibilityIssues, issue => issue.Contains("/@theme") || issue.Contains("/@tint") || issue.Contains("/@indexed"));
        Assert.DoesNotContain(snapshot.FileState.CompatibilityIssues, issue => issue.Contains("主题字体"));
        _storage.Save(PathFor("theme-preserved.xlsq"), snapshot);
        Assert.Equal(File.ReadAllBytes(path), ReadZip(File.ReadAllBytes(PathFor("theme-preserved.xlsq")))["workbook.xlsx"]);
        _storage.Export(PathFor("color-copy.xlsx"), snapshot, false, true);
        Assert.Equal(style, _storage.Open(PathFor("color-copy.xlsx")).Styles[0]);
        var sourceParts = ReadZip(File.ReadAllBytes(path));
        var savedParts = ReadZip(File.ReadAllBytes(PathFor("color-copy.xlsx")));
        Assert.Equal(sourceParts["xl/theme/theme1.xml"], savedParts["xl/theme/theme1.xml"]);
    }

    [Theory]
    [InlineData("theme=12")]
    [InlineData("theme=-1")]
    [InlineData("indexed=66")]
    [InlineData("rgb=ZZ112233")]
    [InlineData("theme=5;tint=NaN")]
    [InlineData("theme=5;tint=1.1")]
    public void InvalidColorReferencesAndTintsAreRejected(string attributes) =>
        Assert.Throws<InvalidDataException>(() => _storage.Open(ColorFixture(attributes)));

    [Fact]
    public void CustomIndexedPaletteAndMissingThemeDoNotSilentlyUseWrongColors()
    {
        var path = ColorFixture("indexed=0"); var parts = ReadZip(File.ReadAllBytes(path));
        var styles = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/styles.xml"]));
        styles.Root!.Add(new XElement(Ns + "colors", new XElement(Ns + "indexedColors",
            new XElement(Ns + "rgbColor", new XAttribute("rgb", "00123456")))));
        parts["xl/styles.xml"] = Encoding.UTF8.GetBytes(styles.ToString()); File.WriteAllBytes(path, Zip(parts));
        Assert.Equal("FF123456", _storage.Open(path).Styles[0].FontColor);
        path = ColorFixture("theme=5", includeTheme: false);
        var loaded = _storage.Open(path);
        Assert.Null(loaded.Styles[0].FontColor);
        Assert.Contains(loaded.FileState!.CompatibilityIssues, issue => issue.Contains("主题颜色 5"));
    }

    private string ColorFixture(string attributes, bool includeTheme = true)
    {
        var path = PathFor("colors.xlsx"); _storage.Export(path, Model([Literal(CellValue.FromText("主题文字"))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var styles = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/styles.xml"]));
        XElement Color(string name) => new(Ns + name, attributes.Split(';').Select(a => { var pair = a.Split('='); return new XAttribute(pair[0], pair[1]); }));
        styles.Descendants(Ns + "font").First().Add(Color("color"));
        // The generated xf references fill 2, after the two reserved fills.
        var pattern = styles.Descendants(Ns + "fill").ElementAt(2).Element(Ns + "patternFill")!;
        pattern.SetAttributeValue("patternType", "solid"); pattern.Add(Color("fgColor"));
        var border = styles.Descendants(Ns + "bottom").First(); border.SetAttributeValue("style", "thin"); border.Add(Color("color"));
        parts["xl/styles.xml"] = Encoding.UTF8.GetBytes(styles.ToString());
        if (includeTheme)
        {
            var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/_rels/workbook.xml.rels"]));
            rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship", new XAttribute("Id", "theme"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"), new XAttribute("Target", "theme/theme1.xml")));
            parts["xl/_rels/workbook.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
            var types = XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/xl/theme/theme1.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")));
            parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(types.ToString());
            parts["xl/theme/theme1.xml"] = Encoding.UTF8.GetBytes("""
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">
                <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
                <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
                <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
                <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
                <a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
                <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
                </a:clrScheme></a:themeElements></a:theme>
                """);
        }
        File.WriteAllBytes(path, Zip(parts)); return path;
    }

    [Theory]
    [InlineData("major", "Cambria", "宋体")]
    [InlineData("minor", "Calibri", "等线")]
    public void ThemeFontsResolveAndSurviveStyleEditingWhileExplicitFontDisconnectsScheme(string scheme, string latin, string eastAsian)
    {
        var path = ColorFixture("theme=1"); var parts = ReadZip(File.ReadAllBytes(path));
        var styles = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/styles.xml"]));
        styles.Descendants(Ns + "font").First().Add(new XElement(Ns + "scheme", new XAttribute("val", scheme)));
        parts["xl/styles.xml"] = Encoding.UTF8.GetBytes(styles.ToString());
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var theme = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/theme/theme1.xml"]));
        theme.Root!.Element(a + "themeElements")!.Add(new XElement(a + "fontScheme", new XAttribute("name", "Test"),
            new XElement(a + "majorFont", new XElement(a + "latin", new XAttribute("typeface", "Cambria")), new XElement(a + "ea", new XAttribute("typeface", "宋体"))),
            new XElement(a + "minorFont", new XElement(a + "latin", new XAttribute("typeface", "Calibri")), new XElement(a + "ea", new XAttribute("typeface", "")),
                new XElement(a + "font", new XAttribute("script", "Hans"), new XAttribute("typeface", "等线")))));
        parts["xl/theme/theme1.xml"] = Encoding.UTF8.GetBytes(theme.ToString()); File.WriteAllBytes(path, Zip(parts));
        var book = _storage.Open(path);
        Assert.Equal(latin, book.Styles[0].FontName); Assert.Equal(eastAsian, book.Styles[0].EastAsianFontName);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellStyleCommand([new(book.Sheets[0].Id, 0, 0)], new(Bold: true)));
        var output = PathFor("theme-edited.xlsq"); _storage.Save(output, session.Snapshot);
        Assert.Equal(session.Snapshot.Styles.ToArray(), _storage.Open(output).Styles.ToArray());
        session.Execute(new SetCellStyleCommand([new(book.Sheets[0].Id, 0, 0)], new(FontName: "Arial")));
        var explicitStyle = session.Snapshot.Styles[^1]; Assert.Null(explicitStyle.FontScheme); Assert.Null(explicitStyle.EastAsianFontName);
        _storage.Save(output, session.Snapshot); Assert.Equal(explicitStyle, _storage.Open(output).Styles[^1]);
    }
}
