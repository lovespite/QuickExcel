using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

/// <summary>MS-XLSX _error/errorType=13. Other rich values and descriptive properties remain diagnosed.</summary>
internal static class XlsxRichErrors
{
    private static readonly XNamespace Rich = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private const string Relation = "http://schemas.microsoft.com/office/2017/06/relationships/";
    private const string Extension = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
    private static XElement[] Items(XElement? parent, XName name)
    {
        if (parent is null) return [];
        var items = parent.Elements(name).ToArray();
        if (items.Length > WorkbookLimits.MaxCells || parent.Attribute("count") is { } count && Int(count.Value, WorkbookLimits.MaxCells) != items.Length)
            throw new InvalidDataException("富值元数据数量不一致或超限。");
        return items;
    }
    public static IReadOnlyDictionary<int, CellErrorCode?> Read(XlsxPackage package, string workbook)
    {
        var result = new Dictionary<int, CellErrorCode?>();
        var path = package.RelatedByType(workbook, "sheetMetadata");
        if (path is null) return result;
        package.RequireType(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
        var metadata = package.XmlPart(path).Root!;
        if (metadata.Name != Main + "metadata") throw new InvalidDataException("无效元数据部件。");
        package.Shape(metadata, "metadataTypes futureMetadata valueMetadata", "", path);
        foreach (var name in new[] { "metadataTypes", "valueMetadata" })
            if (metadata.Elements(Main + name).Count() > 1) throw new InvalidDataException("重复的富值元数据集合。");
        var records = Items(metadata.Element(Main + "valueMetadata"), Main + "bk");
        for (var i = 0; i < records.Length; i++) result[i + 1] = null;
        var dataPath = package.RelatedByUri(workbook, Relation + "rdRichValue");
        var structurePath = package.RelatedByUri(workbook, Relation + "rdRichValueStructure");
        if (dataPath is null || structurePath is null)
        { package.Loss("此元数据尚未建模。", "rich-values/metadata", path); return result; }
        package.RequireType(dataPath, "application/vnd.ms-excel.rdrichvalue+xml");
        package.RequireType(structurePath, "application/vnd.ms-excel.rdrichvaluestructure+xml");
        var data = package.XmlPart(dataPath).Root!; var structuresRoot = package.XmlPart(structurePath).Root!;
        if (data.Name != Rich + "rvData" || structuresRoot.Name != Rich + "rvStructures") throw new InvalidDataException("无效富值数据部件。");
        package.Shape(data, "rv", "count", dataPath); package.Shape(structuresRoot, "s", "count", structurePath);
        var structures = Items(structuresRoot, Rich + "s"); var values = Items(data, Rich + "rv"); var errors = new CellErrorCode?[values.Length];
        for (var i = 0; i < values.Length; i++)
        {
            package.Budget.Check(); var item = values[i]; package.Shape(item, "v", "s", dataPath);
            var structure = structures[Int(Attr(item, "s"), structures.Length - 1)]; package.Shape(structure, "k", "t", structurePath);
            var keys = structure.Elements(Rich + "k").ToArray(); var cells = item.Elements(Rich + "v").ToArray();
            if (keys.Length != cells.Length || keys.Select(k => Attr(k, "n")).Distinct().Count() != keys.Length) throw new InvalidDataException("富值属性数量或名称无效。");
            for (var k = 0; k < keys.Length; k++) { package.Shape(keys[k], "", "n t", structurePath); package.Shape(cells[k], "", "", dataPath); }
            var errorKey = Array.FindIndex(keys, k => (string?)k.Attribute("n") == "errorType" && (string?)k.Attribute("t") == "i");
            if ((string?)structure.Attribute("t") == "_error" && errorKey >= 0 && cells[errorKey].Value == "13")
            {
                errors[i] = CellErrorCode.Calculation;
                if (keys.Length != 1) package.Loss("#CALC! 附加说明属性未建模；修改后的副本仅保留错误身份。", "rich-error/details", dataPath);
            }
            else package.Loss("富值类型尚未建模。", "rich-values/type", dataPath);
        }
        var metadataTypes = Items(metadata.Element(Main + "metadataTypes"), Main + "metadataType");
        if (metadataTypes.Select(t => Attr(t, "name")).Distinct().Count() != metadataTypes.Length) throw new InvalidDataException("元数据类型名称重复。");
        var richType = Array.FindIndex(metadataTypes, t => (string?)t.Attribute("name") == "XLRICHVALUE") + 1;
        foreach (var type in metadataTypes)
        {
            package.Shape(type, "", "name minSupportedVersion copy pasteAll pasteValues merge splitFirst rowColShift clearFormats clearComments assign coerce", path);
            if ((string?)type.Attribute("name") != "XLRICHVALUE") package.Loss("元数据类型尚未建模。", "rich-values/metadata-type", path);
            else if ((string?)type.Attribute("minSupportedVersion") != "120000" ||
                new[] { "copy", "pasteAll", "pasteValues", "merge", "splitFirst", "rowColShift", "clearFormats", "clearComments", "assign", "coerce" }.Any(flag => !Bool((string?)type.Attribute(flag))))
                package.Loss("富值元数据操作策略尚未建模。", "rich-values/policy", path);
        }
        var future = metadata.Elements(Main + "futureMetadata").Where(e => (string?)e.Attribute("name") == "XLRICHVALUE").ToArray();
        if (future.Length != 1 || richType == 0) throw new InvalidDataException("富值元数据绑定缺失或重复。");
        foreach (var other in metadata.Elements(Main + "futureMetadata").Where(e => (string?)e.Attribute("name") != "XLRICHVALUE"))
            package.Loss("未来元数据类型尚未建模。", "rich-values/future-metadata", path, other);
        package.Shape(future[0], "bk", "name count", path);
        var bindings = Items(future[0], Main + "bk"); var indexes = new int[bindings.Length];
        for (var i = 0; i < bindings.Length; i++)
        {
            var binding = bindings[i]; package.Shape(binding, "extLst", "", path);
            var extensions = binding.Elements(Main + "extLst").Elements(Main + "ext").ToArray();
            var matching = extensions.Where(e => (string?)e.Attribute("uri") == Extension).ToArray();
            if (matching.Length != 1 || matching[0].Elements(Rich + "rvb").Count() != 1) throw new InvalidDataException("富值扩展绑定无效。");
            if (extensions.Length != 1 || matching[0].Elements().Count() != 1) package.Loss("富值扩展有未建模内容。", "rich-values/binding", path);
            var reference = matching[0].Element(Rich + "rvb")!; package.Shape(reference, "", "i", path);
            indexes[i] = Int(Attr(reference, "i"), errors.Length - 1);
        }
        for (var i = 0; i < records.Length; i++)
        {
            var record = records[i]; package.Shape(record, "rc", "", path); var references = record.Elements(Main + "rc").ToArray();
            if (references.Length != 1) { package.Loss("复合富值元数据未建模。", "rich-values/record", path); continue; }
            var reference = references[0]; package.Shape(reference, "", "t v", path);
            if (Int(Attr(reference, "t"), metadataTypes.Length) != richType) continue;
            result[i + 1] = errors[indexes[Int(Attr(reference, "v"), indexes.Length - 1)]];
        }
        return result;
    }
    public static CellValue Apply(XElement cell, CellValue value, IReadOnlyDictionary<int, CellErrorCode?> errors, XlsxPackage package, string path)
    {
        if (cell.Attribute("vm") is not { } attribute || attribute.Value == "0") return value;
        var index = Int(attribute.Value, WorkbookLimits.MaxCells);
        if (!errors.TryGetValue(index, out var code)) throw new InvalidDataException("单元格富值元数据引用越界。");
        if (code is null) { package.Loss("单元格富值身份未建模，保留普通回退值。", "cell/value-metadata", path, cell); return value; }
        if ((string?)cell.Attribute("t") != "e" || value != CellValue.FromError(CellErrorCode.Value)) throw new InvalidDataException("#CALC! 富值与回退错误不一致。");
        return CellValue.FromError(code.Value);
    }
    public static void Write(WorkbookSnapshot snapshot, Dictionary<string, byte[]> parts, XElement relationships, XElement types)
    {
        if (!snapshot.Sheets.Any(s => s.Cells.Values.Any(c => c.Value is { Kind: CellValueKind.Error, ErrorCode: CellErrorCode.Calculation }))) return;
        void Part(string path, string type, string relation, XElement root)
        {
            parts["xl/" + path] = XmlBytes(new XDocument(root));
            types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/xl/" + path), new XAttribute("ContentType", type)));
            relationships.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "rich-error-" + parts.Count), new XAttribute("Type", relation), new XAttribute("Target", path)));
        }
        var type = new XElement(Main + "metadataType", new XAttribute("name", "XLRICHVALUE"), new XAttribute("minSupportedVersion", "120000"));
        foreach (var flag in new[] { "copy", "pasteAll", "pasteValues", "merge", "splitFirst", "rowColShift", "clearFormats", "clearComments", "assign", "coerce" }) type.Add(new XAttribute(flag, "1"));
        Part("metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml", Rel.NamespaceName + "/sheetMetadata",
            new(Main + "metadata", new XElement(Main + "metadataTypes", new XAttribute("count", 1), type),
                new XElement(Main + "futureMetadata", new XAttribute("name", "XLRICHVALUE"), new XAttribute("count", 1),
                    new XElement(Main + "bk", new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", Extension), new XElement(Rich + "rvb", new XAttribute("i", 0)))))),
                new XElement(Main + "valueMetadata", new XAttribute("count", 1), new XElement(Main + "bk", new XElement(Main + "rc", new XAttribute("t", 1), new XAttribute("v", 0))))));
        Part("richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml", Relation + "rdRichValue",
            new(Rich + "rvData", new XAttribute("count", 1), new XElement(Rich + "rv", new XAttribute("s", 0), new XElement(Rich + "v", 13))));
        Part("richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml", Relation + "rdRichValueStructure",
            new(Rich + "rvStructures", new XAttribute("count", 1), new XElement(Rich + "s", new XAttribute("t", "_error"), new XElement(Rich + "k", new XAttribute("n", "errorType"), new XAttribute("t", "i")))));
    }
}
