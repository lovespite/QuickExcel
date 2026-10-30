using System.Collections.Immutable;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.TypedWorkbookXml;

namespace QuickExcel.Storage;

internal static class XlsxPrinting
{
    private const string PrinterType = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
    private static string AttributeName(PropertyInfo property) => char.ToLowerInvariant(property.Name[0]) + property.Name[1..];
    private static T? Scalars<T>(XlsxPackage package, string path, XElement? element, T value) where T : class
    {
        if (element is null) return null;
        var properties = typeof(T).GetProperties();
        package.Shape(element, "", string.Join(' ', properties.Select(AttributeName)), path);
        foreach (var property in properties)
        {
            if (element.Attribute(AttributeName(property)) is not { } attribute) continue;
            var type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            object scalar = type == typeof(bool) ? Bool(attribute.Value) : type == typeof(uint) ? uint.Parse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture) : attribute.Value;
            property.SetValue(value, scalar);
        }
        return value;
    }
    private static XElement? WriteScalars(string name, object? value)
    {
        if (value is null) return null;
        return new(Main + name, value.GetType().GetProperties().Select(p => p.GetValue(value) is { } v ?
            new XAttribute(AttributeName(p), v is bool boolean ? boolean ? "1" : "0" : v) : null));
    }
    public static WorksheetPrinting Read(XlsxPackage package, string path, XElement root)
    {
        XElement? Single(string name)
        {
            if (root.Elements(Main + name).Count() > 1) throw new InvalidDataException("重复的打印设置：" + name);
            return root.Element(Main + name);
        }
        var margins = Single("pageMargins");
        PageMargins? marginValues = null;
        if (margins is not null)
        {
            package.Shape(margins, "", "left right top bottom header footer", path);
            double Number(string name) => double.Parse(XlsxPackage.Attr(margins, name), CultureInfo.InvariantCulture);
            marginValues = new(Number("left"), Number("right"), Number("top"), Number("bottom"), Number("header"), Number("footer"));
        }
        var setup = Single("pageSetup");
        var printer = ImmutableArray<byte>.Empty;
        if (setup?.Attribute(Rel + "id") is { } relation)
        {
            var printerPath = package.Related(path, relation.Value, "printerSettings");
            package.RequireType(printerPath, PrinterType);
            if (package.Relationships(printerPath).Count != 0)
                package.Loss("打印机配置包含额外关系，仅原包保留。", "printerSettings/relationships", printerPath);
            else
            {
                if (package.Parts[printerPath].Length > 1024 * 1024) throw new InvalidDataException("打印机配置超过 1 MiB。");
                printer = package.Parts[printerPath].ToImmutableArray();
            }
            setup = new XElement(setup); setup.Attribute(Rel + "id")!.Remove();
        }
        var header = Single("headerFooter");
        PageHeaderFooter? headerValues = null;
        if (header is not null)
        {
            package.Shape(header, "oddHeader oddFooter evenHeader evenFooter firstHeader firstFooter", "differentOddEven differentFirst scaleWithDoc alignWithMargins", path);
            bool? Flag(string name) => header.Attribute(name) is { } a ? Bool(a.Value) : null;
            string? Text(string name)
            {
                if (header.Elements(Main + name).Count() > 1) throw new InvalidDataException("重复页眉页脚。");
                if (header.Element(Main + name) is not { } e) return null;
                package.Shape(e, "", "", path); return Decode(e.Value);
            }
            headerValues = new() { DifferentOddEven = Flag("differentOddEven"), DifferentFirst = Flag("differentFirst"), ScaleWithDoc = Flag("scaleWithDoc"), AlignWithMargins = Flag("alignWithMargins"),
                OddHeader = Text("oddHeader"), OddFooter = Text("oddFooter"), EvenHeader = Text("evenHeader"), EvenFooter = Text("evenFooter"), FirstHeader = Text("firstHeader"), FirstFooter = Text("firstFooter") };
        }
        ImmutableArray<PrintPageBreak> Breaks(string name)
        {
            if (Single(name) is not { } breaks) return [];
            package.Shape(breaks, "brk", "count manualBreakCount", path);
            if (breaks.Elements(Main + "brk").Count() > 1026) throw new InvalidDataException("分页数量超过 1026。");
            var values = breaks.Elements(Main + "brk").Select(e =>
            {
                package.Shape(e, "", "id min max man pt", path);
                return new PrintPageBreak(Int(XlsxPackage.Attr(e, "id"), PrintRange.MaxRows), Int((string?)e.Attribute("min") ?? "0", PrintRange.MaxRows),
                    Int((string?)e.Attribute("max") ?? "0", PrintRange.MaxRows), Bool((string?)e.Attribute("man")), Bool((string?)e.Attribute("pt")));
            }).ToImmutableArray();
            if (breaks.Attribute("count") is { } count && Int(count.Value, 1026) != values.Length ||
                breaks.Attribute("manualBreakCount") is { } manual && Int(manual.Value, 1026) != values.Count(b => b.Manual))
                throw new InvalidDataException("分页计数不一致。");
            return values;
        }
        if (root.Element(Main + "sheetPr")?.Elements(Main + "pageSetUpPr").Count() > 1) throw new InvalidDataException("重复页面设置属性。");
        var result = new WorksheetPrinting { Margins = marginValues, Setup = Scalars(package, path, setup, new PageSetup()),
            Properties = Scalars(package, path, root.Element(Main + "sheetPr")?.Element(Main + "pageSetUpPr"), new PageSetupProperties()),
            Options = Scalars(package, path, Single("printOptions"), new PrintOptions()), HeaderFooter = headerValues,
            RowBreaks = Breaks("rowBreaks"), ColumnBreaks = Breaks("colBreaks"), PrinterSettings = printer };
        _ = new WorksheetSnapshot(SheetId.New(), "Print", printing: result);
        return result;
    }

    public static WorksheetPrinting ReadName(WorksheetSnapshot sheet, string value, bool titles)
    {
        // Commas and apostrophes inside quoted sheet names are not union separators.
        var prefix = @"(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!,]+))!";
        var cell = @"\$(?<c1>[A-Z]{1,3})\$(?<r1>[1-9][0-9]{0,6})(?::\$(?<c2>[A-Z]{1,3})\$(?<r2>[1-9][0-9]{0,6}))?";
        var rows = @"\$(?<r1>[1-9][0-9]{0,6}):\$(?<r2>[1-9][0-9]{0,6})";
        var cols = @"\$(?<c1>[A-Z]{1,3}):\$(?<c2>[A-Z]{1,3})";
        var pattern = new Regex(@"\G" + prefix + "(?:" + cell + "|" + rows + "|" + cols + @")(?:,|$)", RegexOptions.CultureInvariant);
        var ranges = ImmutableArray.CreateBuilder<PrintRange>(); var offset = 0;
        while (offset < value.Length)
        {
            var match = pattern.Match(value, offset);
            if (!match.Success || match.Index != offset) throw new NotSupportedException("打印名称不是受支持的本表绝对区域/整行整列引用。");
            var name = match.Groups["quoted"].Success ? match.Groups["quoted"].Value.Replace("''", "'") : match.Groups["plain"].Value;
            if (!name.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase)) throw new NotSupportedException("打印名称引用其他工作表。");
            int Row(string key) => int.Parse(match.Groups[key].Value, CultureInfo.InvariantCulture) - 1;
            int Col(string key) { var n = 0; foreach (var c in match.Groups[key].Value) n = n * 26 + c - 'A' + 1; return n - 1; }
            var kind = !match.Groups["c1"].Success ? PrintRangeKind.Rows : !match.Groups["r1"].Success ? PrintRangeKind.Columns : PrintRangeKind.Cells;
            var firstRow = kind == PrintRangeKind.Columns ? 0 : Row("r1"); var firstColumn = kind == PrintRangeKind.Rows ? 0 : Col("c1");
            ranges.Add(new(kind, firstRow, firstColumn, kind == PrintRangeKind.Columns ? PrintRange.MaxRows - 1 : match.Groups["r2"].Success ? Row("r2") : firstRow,
                kind == PrintRangeKind.Rows ? PrintRange.MaxColumns - 1 : match.Groups["c2"].Success ? Col("c2") : firstColumn));
            if (ranges.Count > 1024) throw new InvalidDataException("打印区域数量过多。");
            offset += match.Length;
        }
        if (ranges.Count == 0 || value.EndsWith(',')) throw new NotSupportedException("打印名称缺少区域。");
        var printing = titles ? sheet.Printing with { Titles = ranges.ToImmutable() } : sheet.Printing with { Areas = ranges.ToImmutable() };
        _ = sheet.WithPrinting(printing); return printing;
    }
    public static IEnumerable<XElement> WriteNames(WorksheetSnapshot sheet, int index)
    {
        string Column(int column) { var text = ""; for (var c = column + 1; c > 0; c = (c - 1) / 26) text = (char)('A' + (c - 1) % 26) + text; return "$" + text; }
        string Range(PrintRange r)
        {
            var reference = r.Kind switch { PrintRangeKind.Rows => $"${r.Top + 1}:${r.Bottom + 1}", PrintRangeKind.Columns => $"{Column(r.Left)}:{Column(r.Right)}",
                _ => $"{Column(r.Left)}${r.Top + 1}:{Column(r.Right)}${r.Bottom + 1}" };
            return "'" + sheet.Name.Replace("'", "''") + "'!" + reference;
        }
        foreach (var (name, ranges) in new[] { ("_xlnm.Print_Area", sheet.Printing.Areas), ("_xlnm.Print_Titles", sheet.Printing.Titles) })
            if (!ranges.IsEmpty) yield return new(Main + "definedName", new XAttribute("name", name), new XAttribute("localSheetId", index), Encode(string.Join(',', ranges.Select(Range))));
    }
    public static XElement? WriteProperties(WorksheetPrinting printing) => WriteScalars("pageSetUpPr", printing.Properties);
    public static void WriteSheet(WorksheetPrinting printing, XElement root)
    {
        root.Add(WriteScalars("printOptions", printing.Options), WriteScalars("pageMargins", printing.Margins), WriteScalars("pageSetup", printing.Setup));
        if (!printing.PrinterSettings.IsEmpty) root.Element(Main + "pageSetup")!.SetAttributeValue(Rel + "id", "printer");
        if (printing.HeaderFooter is { } header)
        {
            var element = new XElement(Main + "headerFooter");
            foreach (var property in typeof(PageHeaderFooter).GetProperties())
                if (property.GetValue(header) is { } value)
                {
                    var name = AttributeName(property);
                    if (value is bool flag) element.SetAttributeValue(name, flag ? 1 : 0);
                    else element.Add(new XElement(Main + name, Encode((string)value)));
                }
            root.Add(element);
        }
        foreach (var (name, breaks) in new[] { ("rowBreaks", printing.RowBreaks), ("colBreaks", printing.ColumnBreaks) })
            if (!breaks.IsEmpty) root.Add(new XElement(Main + name, new XAttribute("count", breaks.Length), new XAttribute("manualBreakCount", breaks.Count(b => b.Manual)),
                breaks.Select(b => new XElement(Main + "brk", new XAttribute("id", b.Id), new XAttribute("min", b.Minimum), new XAttribute("max", b.Maximum),
                    new XAttribute("man", b.Manual ? 1 : 0), new XAttribute("pt", b.Pivot ? 1 : 0)))));
    }
    public static void WritePrinter(WorksheetPrinting printing, int sheet, Dictionary<string, byte[]> parts, XElement types)
    {
        if (printing.PrinterSettings.IsEmpty) return;
        var path = $"xl/printerSettings/printerSettings{sheet}.bin";
        parts[path] = printing.PrinterSettings.ToArray();
        types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", PrinterType)));
        var relPath = RelationshipPart($"xl/worksheets/sheet{sheet}.xml");
        var relations = parts.TryGetValue(relPath, out var existing) ? Xml(existing, new(CancellationToken.None)).Root! : new XElement(PackageRel + "Relationships");
        relations.Add(Relationship("printer", "printerSettings", $"../printerSettings/printerSettings{sheet}.bin"));
        parts[relPath] = XmlBytes(new XDocument(relations));
    }
}
