using System.Collections.Immutable;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;
using static QuickExcel.Storage.TypedWorkbookXml;

namespace QuickExcel.Storage;

internal static class XlsxTables
{
    private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
    internal static void Shape(XElement element, string children, string attributes)
    {
        var childNames = children.Split(' ', StringSplitOptions.RemoveEmptyEntries); var attributeNames = attributes.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (element.Elements().Any(e => e.Name.Namespace != Main || !childNames.Contains(e.Name.LocalName)) || element.Attributes().Any(a => !a.IsNamespaceDeclaration &&
            !(a.Name.NamespaceName.Length == 0 && attributeNames.Contains(a.Name.LocalName)) &&
            !(a.Name.LocalName == "uid" && a.Name.NamespaceName is "http://schemas.microsoft.com/office/spreadsheetml/2014/revision" or "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" or "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9" && Guid.TryParse(a.Value, out _)) &&
            a.Name != XName.Get("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006")))
            throw new NotSupportedException("未支持的表对象属性或子元素：" + element.Name.LocalName);
    }
    public static ImmutableArray<WorksheetTable> Read(XlsxPackage package, string workbook, string sheetPath, XElement sheet, List<PreservedFormula> formulas, string sheetName)
    {
        var tables = ImmutableArray.CreateBuilder<WorksheetTable>(); var paths = new HashSet<string>();
        if (sheet.Elements(Main + "tableParts").Count() > 1) throw new InvalidDataException("表部件列表重复。");
        if (sheet.Element(Main + "tableParts") is not { } parts) return [];
        Shape(parts, "tablePart", "count");
        if (parts.Attribute("count") is { } count && Int(count.Value, 256) != parts.Elements().Count()) throw new InvalidDataException("表部件数量不一致。");
        foreach (var entry in parts.Elements())
        {
            if (tables.Count >= 256) throw new InvalidDataException("表数量超过 256。");
            if (entry.HasElements || entry.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != Rel + "id")) throw new InvalidDataException("表部件关系无效。");
            var path = package.Related(sheetPath, Attr(entry, Rel + "id"), "table");
            if (!paths.Add(path)) throw new InvalidDataException("表部件关系重复。");
            if (paths.Count > 256) throw new InvalidDataException("表部件数量超过 256。");
            package.RequireType(path, ContentType); var root = package.XmlPart(path).Root!;
            if (root.Name != Main + "table") throw new InvalidDataException("无效表部件。");
            try
            {
                Shape(root, "autoFilter sortState tableColumns tableStyleInfo", "id name displayName ref headerRowCount totalsRowCount totalsRowShown");
                if (root.Elements().GroupBy(e => e.Name).Any(g => g.Count() > 1)) throw new InvalidDataException("表对象子元素重复。");
                var name = Decode(Attr(root, "displayName"));
                if (root.Attribute("name") is { } storedName && Decode(storedName.Value) != name) throw new NotSupportedException("表 name 与 displayName 不同，尚不支持受控编辑。");
                var range = CellRectangle.Parse(Attr(root, "ref"));
                var header = root.Attribute("headerRowCount") is not { } headerCount || Int(headerCount.Value, 1) == 1;
                var total = root.Attribute("totalsRowCount") is { } totalCount && Int(totalCount.Value, 1) == 1;
                var shown = root.Attribute("totalsRowShown") is not { } shownValue || Bool(shownValue.Value);
                var columns = root.Element(Main + "tableColumns") ?? throw new InvalidDataException("表列定义缺失。");
                Shape(columns, "tableColumn", "count"); var list = new List<TableColumn>();
                if (Int(Attr(columns, "count"), WorkbookLimits.MaxColumns) != columns.Elements().Count()) throw new InvalidDataException("表列数量不一致。");
                foreach (var column in columns.Elements())
                {
                    Shape(column, "calculatedColumnFormula totalsRowFormula", "id name totalsRowFunction totalsRowLabel dataDxfId headerRowDxfId totalsRowDxfId");
                    if (column.Elements().GroupBy(e => e.Name).Any(g => g.Count() > 1)) throw new InvalidDataException("表列公式重复。");
                    string? Formula(string field)
                    {
                        if (column.Element(Main + field) is not { } value) return null;
                        Shape(value, "", ""); if (value.Value.Length == 0) throw new InvalidDataException("表公式为空。"); return "=" + Decode(value.Value).TrimStart('=');
                    }
                    list.Add(new(Int(Attr(column, "id"), int.MaxValue), Decode(Attr(column, "name")), (string?)column.Attribute("totalsRowFunction"),
                        column.Attribute("totalsRowLabel") is { } label ? Decode(label.Value) : null, Formula("calculatedColumnFormula"), Formula("totalsRowFormula"))
                    {
                        DataStyle = XlsxTableStyles.ColumnStyle(package, workbook, column.Attribute("dataDxfId")),
                        HeaderStyle = XlsxTableStyles.ColumnStyle(package, workbook, column.Attribute("headerRowDxfId")),
                        TotalsStyle = XlsxTableStyles.ColumnStyle(package, workbook, column.Attribute("totalsRowDxfId"))
                    });
                }
                var info = root.Element(Main + "tableStyleInfo"); TableStyleInfo style = new();
                if (info is not null)
                {
                    Shape(info, "", "name showFirstColumn showLastColumn showRowStripes showColumnStripes");
                    var styleName = (string?)info.Attribute("name");
                    style = new(styleName, Bool((string?)info.Attribute("showFirstColumn")), Bool((string?)info.Attribute("showLastColumn")),
                        Bool((string?)info.Attribute("showRowStripes")), Bool((string?)info.Attribute("showColumnStripes"))) { Elements = XlsxTableStyles.Resolve(package, workbook, styleName) };
                }
                var data = XlsxFiltering.Read(package, workbook, path, root);
                var table = new WorksheetTable(Int(Attr(root, "id"), int.MaxValue), name, range, list, header, total, shown, style, data.Filter, data.Sort);
                foreach (var column in list.Where(c => c.CalculatedFormula is not null || c.TotalsFormula is not null))
                    formulas.Add(new(PreservedFormulaKind.StructuredReference, path, sheetName, range.ToString(), range.ToString(), column.CalculatedFormula ?? column.TotalsFormula!));
                tables.Add(table);
            }
            catch (NotSupportedException ex) { package.Loss(ex.Message + " 表对象只随原包保留。", "table", path, root); }
        }
        return tables.ToImmutable();
    }
    public static void Write(WorksheetSnapshot sheet, int index, Dictionary<string, byte[]> parts, XElement root, XElement types, XElement styles, bool valuesOnly)
    {
        if (sheet.Data.Tables.IsEmpty) return;
        var path = $"xl/worksheets/sheet{index}.xml"; var relPath = RelationshipPart(path);
        var relations = parts.TryGetValue(relPath, out var bytes) ? Xml(bytes, new(CancellationToken.None)).Root! : new XElement(PackageRel + "Relationships");
        var tableParts = new XElement(Main + "tableParts", new XAttribute("count", sheet.Data.Tables.Length));
        foreach (var table in sheet.Data.Tables)
        {
            var tablePath = $"xl/tables/table{table.Id}.xml"; var relationship = "table" + table.Id;
            relations.Add(Relationship(relationship, "table", $"../tables/table{table.Id}.xml")); tableParts.Add(new XElement(Main + "tablePart", new XAttribute(Rel + "id", relationship)));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + tablePath), new XAttribute("ContentType", ContentType)));
            var xml = new XElement(Main + "table", new XAttribute("id", table.Id), new XAttribute("name", Encode(table.Name)), new XAttribute("displayName", Encode(table.Name)), new XAttribute("ref", table.Range),
                new XAttribute("headerRowCount", table.HeaderRow ? 1 : 0), new XAttribute("totalsRowCount", table.TotalsRow ? 1 : 0), new XAttribute("totalsRowShown", table.TotalsRowShown ? 1 : 0));
            var sort = XlsxFiltering.WriteSort(table.Sort, styles); xml.Add(XlsxFiltering.WriteFilter(table.Filter, sort)); if (table.Filter is null) xml.Add(sort);
            xml.Add(new XElement(Main + "tableColumns", new XAttribute("count", table.Columns.Length), table.Columns.Select(column =>
            {
                var element = new XElement(Main + "tableColumn", new XAttribute("id", column.Id), new XAttribute("name", Encode(column.Name)));
                if (!valuesOnly && column.TotalsFunction is { } function) element.Add(new XAttribute("totalsRowFunction", function));
                if (column.TotalsLabel is { } label) element.Add(new XAttribute("totalsRowLabel", Encode(label)));
                foreach (var (attribute, style) in new[] { ("dataDxfId", column.DataStyle), ("headerRowDxfId", column.HeaderStyle), ("totalsRowDxfId", column.TotalsStyle) })
                    if (style is not null) element.Add(new XAttribute(attribute, XlsxTableStyles.WriteDxf(style, styles)));
                if (!valuesOnly && column.CalculatedFormula is { } calculated) element.Add(new XElement(Main + "calculatedColumnFormula", Encode(calculated.TrimStart('='))));
                if (!valuesOnly && column.TotalsFormula is { } formula) element.Add(new XElement(Main + "totalsRowFormula", Encode(formula.TrimStart('='))));
                return element;
            })));
            XlsxTableStyles.WriteCustom(table.Style, styles);
            xml.Add(new XElement(Main + "tableStyleInfo", table.Style.Name is { } name ? new XAttribute("name", name) : null,
                new XAttribute("showFirstColumn", table.Style.FirstColumn ? 1 : 0), new XAttribute("showLastColumn", table.Style.LastColumn ? 1 : 0),
                new XAttribute("showRowStripes", table.Style.RowStripes ? 1 : 0), new XAttribute("showColumnStripes", table.Style.ColumnStripes ? 1 : 0)));
            parts[tablePath] = XmlBytes(new XDocument(xml));
        }
        root.Add(tableParts); parts[relPath] = XmlBytes(new XDocument(relations));
    }
}
