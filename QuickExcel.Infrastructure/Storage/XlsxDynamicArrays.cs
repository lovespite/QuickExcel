using System.Collections.Immutable;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

/// <summary>MS-XLSX XLDAPR metadata. Unknown metadata is retained through the existing package guard.</summary>
internal static class XlsxDynamicArrays
{
    private static readonly XNamespace Dynamic = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
    private const string Extension = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";
    private static readonly string[] Flags = ["copy", "pasteAll", "pasteValues", "merge", "splitFirst", "rowColShift", "clearFormats", "clearComments", "assign", "coerce", "cellMeta"];

    public static IReadOnlySet<int> Read(XlsxPackage package, string workbook)
    {
        var accepted = new HashSet<int>(); var path = package.RelatedByType(workbook, "sheetMetadata");
        if (path is null) return accepted;
        var metadata = package.XmlPart(path).Root!;
        if (metadata.Name != Main + "metadata") throw new InvalidDataException("无效元数据部件。");
        var types = Items(metadata.Element(Main + "metadataTypes"), "metadataType");
        if (types.Select(t => Attr(t,"name")).Distinct().Count() != types.Length) throw new InvalidDataException("元数据类型重复。");
        var type = Array.FindIndex(types, t => (string?)t.Attribute("name") == "XLDAPR");
        if (type < 0) return accepted;
        var policy = types[type];
        package.Shape(policy, "", "name minSupportedVersion " + string.Join(' ', Flags), path);
        if ((string?)policy.Attribute("minSupportedVersion") != "120000" || Flags.Any(f => !Bool((string?)policy.Attribute(f))))
        { package.Loss("动态数组元数据策略尚未建模。", "dynamic-array/policy", path); return accepted; }
        var future = metadata.Elements(Main + "futureMetadata").Where(e => (string?)e.Attribute("name") == "XLDAPR").ToArray();
        if (future.Length != 1 || metadata.Elements(Main + "cellMetadata").Count() > 1) throw new InvalidDataException("动态数组元数据集合缺失或重复。");
        var bindings = Items(future[0], "bk"); var supported = new HashSet<int>();
        package.Shape(future[0], "bk", "name count", path);
        for (var i = 0; i < bindings.Length; i++)
        {
            package.Budget.Check(); var binding = bindings[i]; package.Shape(binding, "extLst", "", path);
            var extensions = binding.Elements(Main + "extLst").Elements(Main + "ext").ToArray();
            var matching = extensions.Where(e => string.Equals((string?)e.Attribute("uri"), Extension, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matching.Length != 1 || matching[0].Elements(Dynamic + "dynamicArrayProperties").Count() != 1)
                throw new InvalidDataException("动态数组属性绑定无效。");
            var properties = matching[0].Element(Dynamic + "dynamicArrayProperties")!;
            package.Shape(properties, "", "fDynamic fCollapsed", path);
            if (extensions.Length != 1 || matching[0].Elements().Count() != 1 || properties.HasElements ||
                !Bool((string?)properties.Attribute("fDynamic")) || Bool((string?)properties.Attribute("fCollapsed")))
            { package.Loss("动态数组扩展属性尚未完整建模。", "dynamic-array/properties", path); continue; }
            supported.Add(i);
        }
        var records = Items(metadata.Element(Main + "cellMetadata"), "bk");
        for (var i = 0; i < records.Length; i++)
        {
            package.Budget.Check(); var record = records[i]; package.Shape(record, "rc", "", path);
            var references = record.Elements(Main + "rc").ToArray();
            if (references.Length != 1) { package.Loss("复合单元格元数据尚未建模。", "dynamic-array/record", path); continue; }
            var reference = references[0]; package.Shape(reference, "", "t v", path);
            var kind = Int(Attr(reference, "t"), types.Length);
            if (kind == 0) throw new InvalidDataException("单元格元数据类型索引为零。");
            if (kind == type + 1 && supported.Contains(Int(Attr(reference, "v"), bindings.Length - 1))) accepted.Add(i + 1);
            else package.Loss("单元格元数据引用未建模。", "dynamic-array/reference", path);
        }
        return accepted;
    }

    private static XElement[] Items(XElement? parent, string name)
    {
        if (parent is null) return [];
        var elements = parent.Elements(Main + name).ToArray();
        if (elements.Length > WorkbookLimits.MaxCells || parent.Attribute("count") is { } count && Int(count.Value, WorkbookLimits.MaxCells) != elements.Length)
            throw new InvalidDataException("动态数组元数据数量不一致或超限。");
        return elements;
    }

    public static WorksheetSnapshot ApplyCaches(WorksheetSnapshot sheet, IReadOnlyDictionary<CellPosition, CellRectangle> ranges,
        HashSet<CellPosition> cached, XlsxPackage package, string path, List<PreservedFormula> preserved)
    {
        var cells = sheet.Cells.ToBuilder(); var owned = new HashSet<CellPosition>();
        foreach (var (anchor, range) in ranges)
        {
            package.Budget.Check();
            if (range.Top != anchor.Row || range.Left != anchor.Column) throw new InvalidDataException("动态数组锚点必须是区域左上角。");
            if (range.Count > FormulaArray.MaxElements || range.Count + owned.Count > WorkbookLimits.MaxCells)
                throw new InvalidDataException("动态数组缓存超过元素预算。");
            var values = new List<CellValue>((int)range.Count); var positions = new List<CellPosition>((int)range.Count); var complete = true;
            for (var row = range.Top; row <= range.Bottom; row++)
            for (var column = range.Left; column <= range.Right; column++)
            {
                package.Budget.Check(); var position = new CellPosition(row, column);
                if (!owned.Add(position)) throw new InvalidDataException("动态数组缓存区域重叠。");
                var cell = cells.GetValueOrDefault(position);
                if (position != anchor && cell?.Content.Kind == CellContentKind.Formula) throw new InvalidDataException("动态数组子格含独立公式。");
                if (!cached.Contains(position) || cell is null || cell.Content.RichText is not null) complete = false;
                positions.Add(position); values.Add(cell?.Value ?? CellValue.Blank);
            }
            if (!complete)
            {
                preserved.Add(new(PreservedFormulaKind.DynamicArray, path, sheet.Name, TypedWorkbookXml.Reference(anchor), range.ToString(), cells[anchor].Content.FormulaText, HasCompleteCache: false));
                continue;
            }
            package.Budget.Parse(range.Count * 64);
            var array = new FormulaArray(range.Rows, range.Columns, values);
            var owner = new CellAddress(sheet.Id, anchor.Row, anchor.Column);
            cells[anchor] = new(cells[anchor].Content, new(array[0,0], CalculationState.Stale, 0, 0, 0, array: array, blockedSpill: cells[anchor].Calculation?.BlockedSpill));
            for (var i = 1; i < positions.Count; i++)
            {
                var position = positions[i];
                cells[position] = new(CellContent.FromLiteral(CellValue.Blank, cells[position].Content.StyleId),
                    new(values[i], CalculationState.Stale, 0, 0, 0), owner);
            }
        }
        return sheet.WithCells(cells.ToImmutable());
    }

    public static bool Present(WorkbookSnapshot snapshot) => snapshot.Sheets.Any(s => s.Cells.Values.Any(c => c.Content.DynamicArray));

    public static void Write(WorkbookSnapshot snapshot, Dictionary<string, byte[]> parts, XElement relationships, XElement types)
    {
        if (!Present(snapshot)) return;
        XElement metadata;
        if (parts.TryGetValue("xl/metadata.xml", out var existing))
        { using var stream = new MemoryStream(existing); metadata = XDocument.Load(stream).Root!; }
        else
        {
            metadata = new(Main + "metadata", new XElement(Main + "metadataTypes", new XAttribute("count", 0)));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/xl/metadata.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")));
            relationships.Add(TypedWorkbookXml.Relationship("dynamic-array-metadata", "sheetMetadata", "metadata.xml"));
        }
        var metadataTypes = metadata.Element(Main + "metadataTypes")!;
        var typeId = metadataTypes.Elements().Count() + 1;
        var type = new XElement(Main + "metadataType", new XAttribute("name", "XLDAPR"), new XAttribute("minSupportedVersion", "120000"));
        foreach (var flag in Flags) type.Add(new XAttribute(flag, 1));
        metadataTypes.Add(type); metadataTypes.SetAttributeValue("count", typeId);
        var future = new XElement(Main + "futureMetadata", new XAttribute("name", "XLDAPR"), new XAttribute("count", 1),
            new XElement(Main + "bk", new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", Extension),
                new XElement(Dynamic + "dynamicArrayProperties", new XAttribute("fDynamic", 1), new XAttribute("fCollapsed", 0))))));
        var cellMetadata = new XElement(Main + "cellMetadata", new XAttribute("count", 1),
            new XElement(Main + "bk", new XElement(Main + "rc", new XAttribute("t", typeId), new XAttribute("v", 0))));
        var valueMetadata = metadata.Element(Main + "valueMetadata");
        if (valueMetadata is not null) valueMetadata.AddBeforeSelf(future, cellMetadata); else metadata.Add(future, cellMetadata);
        parts["xl/metadata.xml"] = XmlBytes(new XDocument(metadata));
    }
}
