using System.Collections.Immutable;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static class XlsxFiltering
{
    private const string Extension = "urn:quickexcel:filter-rows:v1";
    private static readonly XNamespace Quick = Extension;
    private static readonly string[] Operators = ["equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"];
    private static void Shape(XElement element, string children, string attributes)
    {
        var allowedChildren = children.Split(' ', StringSplitOptions.RemoveEmptyEntries); var allowedAttributes = attributes.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (element.Elements().Any(e => e.Name.Namespace != Main || !allowedChildren.Contains(e.Name.LocalName)) ||
            element.Attributes().Any(a => !a.IsNamespaceDeclaration && !(a.Name.NamespaceName.Length == 0 && allowedAttributes.Contains(a.Name.LocalName)) &&
                !(a.Name == XName.Get("uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision") && Guid.TryParse(a.Value, out _))))
            throw new NotSupportedException("包含未支持的 " + element.Name.LocalName + " 条件或属性。");
    }

    public static WorksheetData Read(XlsxPackage package, string workbook, string path, XElement root)
    {
        WorksheetAutoFilter? filter = null; WorksheetSortState? sort = null;
        if (root.Elements(Main + "autoFilter").Count() > 1 || root.Elements(Main + "sortState").Count() > 1)
            throw new InvalidDataException("筛选/排序状态重复。");
        var auto = root.Element(Main + "autoFilter");
        if (auto is not null)
        {
            try
            {
                Shape(auto, "filterColumn sortState", "ref"); var range = CellRectangle.Parse(Attr(auto, "ref"));
                var columns = new Dictionary<int, AutoFilterColumn>();
                foreach (var column in auto.Elements(Main + "filterColumn"))
                {
                    Shape(column, "filters customFilters colorFilter", "colId hiddenButton showButton");
                    var index = Int(Attr(column, "colId"), range.Right - range.Left) + range.Left;
                    var button = !Bool((string?)column.Attribute("hiddenButton")) && (column.Attribute("showButton") is null || Bool((string?)column.Attribute("showButton")));
                    if (column.Elements().Count() > 1) throw new NotSupportedException("多个筛选条件类型不能混用。");
                    AutoFilterColumn value;
                    if (column.Element(Main + "filters") is { } values)
                    {
                        Shape(values, "filter", "blank");
                        var labels = new List<string>();
                        foreach (var item in values.Elements()) { Shape(item, "", "val"); labels.Add(Decode(item.Attribute("val")?.Value ?? throw new InvalidDataException("筛选值缺失。"))); }
                        if (labels.Count == 0 && !Bool((string?)values.Attribute("blank"))) throw new NotSupportedException("空筛选列表无法确定可重应用语义。");
                        value = new(labels, Bool((string?)values.Attribute("blank")), showButton: button);
                    }
                    else if (column.Element(Main + "customFilters") is { } custom)
                    {
                        Shape(custom, "customFilter", "and");
                        var items = custom.Elements().ToArray();
                        if (items.Length is < 1 or > 2) throw new InvalidDataException("自定义筛选须有 1–2 个条件。");
                        AutoFilterCondition Condition(XElement item)
                        {
                            Shape(item, "", "operator val");
                            var op = Array.IndexOf(Operators, (string?)item.Attribute("operator") ?? "equal");
                            if (op < 0) throw new NotSupportedException("未知筛选比较符。");
                            return new((AutoFilterComparison)op, Decode(item.Attribute("val")?.Value ?? throw new InvalidDataException("筛选条件值缺失。")));
                        }
                        try { value = new(first: Condition(items[0]), second: items.Length == 2 ? Condition(items[1]) : null, useAnd: Bool((string?)custom.Attribute("and")), showButton: button); }
                        catch (ArgumentException ex) { throw new NotSupportedException("自定义筛选不能稳定重新应用：" + ex.Message, ex); }
                    }
                    else if (column.Element(Main + "colorFilter") is { } color)
                    {
                        Shape(color, "", "cellColor dxfId");
                        var font = color.Attribute("cellColor") is not null && !Bool((string?)color.Attribute("cellColor"));
                        var stylePath = package.RelatedByType(workbook, "styles", true)!; var styles = package.XmlPart(stylePath).Root!;
                        var dxfs = styles.Element(Main + "dxfs")?.Elements(Main + "dxf").ToArray() ?? [];
                        var dxf = dxfs[Int(Attr(color, "dxfId"), dxfs.Length - 1)];
                        // Excel stores the comparison color in a fill DXF for BOTH fill and font filters.
                        // cellColor selects which cell property is compared to that color.
                        Shape(dxf, "fill", "");
                        var fill = dxf.Element(Main + "fill") ?? throw new NotSupportedException(MessageText.Get("颜色筛选缺少差异填充定义。"));
                        Shape(fill, "patternFill", ""); var pattern = fill.Element(Main + "patternFill") ?? throw new NotSupportedException(MessageText.Get("颜色筛选缺少图案。"));
                        Shape(pattern, "fgColor bgColor", "patternType");
                        var type = (string?)pattern.Attribute("patternType") ?? "none";
                        if (type is not ("none" or "solid")) throw new NotSupportedException(MessageText.Get("图案颜色筛选尚未支持。"));
                        var colorElement = type == "none" ? null : pattern.Element(Main + "fgColor") ?? pattern.Element(Main + "bgColor");
                        if (type == "solid" && colorElement is null) throw new NotSupportedException(MessageText.Get("纯色筛选缺少颜色。"));
                        var rgb = colorElement is null ? null : new XlsxColors(package, workbook, styles, stylePath).Read(colorElement, package, stylePath);
                        try { value = new(showButton: button, color: new(font, rgb)); }
                        catch (ArgumentException ex) { throw new NotSupportedException(MessageText.Get("颜色筛选无法重应用：") + ex.Message, ex); }
                    }
                    else value = new(showButton: button);
                    if (!columns.TryAdd(index, value)) throw new InvalidDataException("重复的筛选列。");
                }
                filter = new(range, columns);
            }
            catch (NotSupportedException ex) { package.Loss(ex.Message + " 原条件只随原包保留，不重新解释。", "worksheet/autoFilter", path, auto); }
        }
        var sorts = root.Elements(Main + "sortState").Concat(auto?.Elements(Main + "sortState") ?? []).ToArray();
        if (sorts.Length > 1) throw new NotSupportedException("同一工作表有多个排序状态。");
        if (sorts.Length == 1)
        {
            try { sort = ReadSort(package, workbook, sorts[0]); }
            catch (NotSupportedException ex) { package.Loss(ex.Message + " 原排序状态随原包保留。", "worksheet/sortState", path, sorts[0]); }
        }
        return new(filter, sort);
    }

    private static WorksheetSortState ReadSort(XlsxPackage package, string workbook, XElement root)
    {
        Shape(root, "sortCondition", "ref columnSort caseSensitive sortMethod");
        if (Bool((string?)root.Attribute("columnSort"))) throw new NotSupportedException("从左到右的列排序尚未实现。");
        var range = CellRectangle.Parse(Attr(root, "ref")); var keys = new List<SortKey>();
        foreach (var item in root.Elements())
        {
            Shape(item, "", "ref descending sortBy customList dxfId");
            var keyRange = CellRectangle.Parse(Attr(item, "ref"));
            if (keyRange.Left != keyRange.Right || keyRange.Left < range.Left || keyRange.Right > range.Right || keyRange.Top != range.Top || keyRange.Bottom != range.Bottom)
                throw new NotSupportedException("排序键不是完整数据列。");
            var basis = (string?)item.Attribute("sortBy") switch { null or "value" => SortBasis.Value, "cellColor" => SortBasis.FillColor, "fontColor" => SortBasis.FontColor, _ => throw new NotSupportedException("图标排序尚未实现。") };
            string? color = null;
            if (basis != SortBasis.Value)
            {
                var stylePath = package.RelatedByType(workbook, "styles", true)!; var styles = package.XmlPart(stylePath).Root!;
                var dxfs = styles.Element(Main + "dxfs")?.Elements(Main + "dxf").ToArray() ?? [];
                var dxf = dxfs[Int(Attr(item, "dxfId"), dxfs.Length - 1)];
                var colorElement = basis == SortBasis.FontColor ? dxf.Element(Main + "font")?.Element(Main + "color") :
                    dxf.Element(Main + "fill")?.Element(Main + "patternFill")?.Element(Main + "fgColor") ?? dxf.Element(Main + "fill")?.Element(Main + "patternFill")?.Element(Main + "bgColor");
                color = new XlsxColors(package, workbook, styles, stylePath).Read(colorElement, package, stylePath);
                if (color is null) throw new NotSupportedException("排序颜色无法解析。");
            }
            else if (item.Attribute("dxfId") is not null) throw new NotSupportedException("值排序不能绑定差异样式。");
            keys.Add(new(keyRange.Left, Bool((string?)item.Attribute("descending")), Bool((string?)root.Attribute("caseSensitive")), basis, color,
                item.Attribute("customList") is { } list ? Decode(list.Value).Split(',') : null));
        }
        try { return new(range, keys, (string?)root.Attribute("sortMethod") ?? "pinYin"); }
        catch (ArgumentException ex) { throw new NotSupportedException("排序状态包含当前不能重新应用的键：" + ex.Message, ex); }
    }

    public static XElement Shape(XElement root)
    {
        var result = new XElement(root);
        result.Elements(Main + "extLst").Elements(Main + "ext").Where(e => (string?)e.Attribute("uri") == Extension).Remove();
        result.Elements(Main + "extLst").Where(e => !e.HasElements && !e.HasAttributes).Remove(); return result;
    }

    private static ImmutableHashSet<int> Rows(string? text)
    {
        if (string.IsNullOrEmpty(text)) return [];
        if (text.Length > 700_000) throw new InvalidDataException("筛选行元数据过大。");
        var result = ImmutableHashSet.CreateBuilder<int>();
        foreach (var token in text.Split(','))
        {
            var row = Int(token, WorkbookLimits.MaxRows) - 1;
            if (row < 0 || !result.Add(row)) throw new InvalidDataException("筛选行元数据重复或越界。");
        }
        return result.ToImmutable();
    }

    public static WorksheetSnapshot ReadHidden(XlsxPackage package, string path, XElement xml, WorksheetSnapshot sheet, IReadOnlyList<StyleDefinition> styles, ExcelDateSystem system)
    {
        var extensions = xml.Elements(Main + "extLst").Elements(Main + "ext").Where(e => (string?)e.Attribute("uri") == Extension).ToArray();
        XElement? metadata = null;
        if (extensions.Length > 0)
        {
            if (extensions.Length != 1 || extensions[0].Elements().Count() != 1 || extensions[0].Elements().Single().Name != Quick + "rows") throw new InvalidDataException("筛选行扩展无效。");
            metadata = extensions[0].Elements().Single();
            if (extensions[0].Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "uri") || metadata.HasElements ||
                metadata.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.NamespaceName.Length != 0 || a.Name.LocalName is not ("filtered" or "manual" or "generated"))))
                throw new InvalidDataException("筛选行扩展字段无效。");
        }
        if (!sheet.Data.AllFilters.Any()) return sheet;
        WorksheetSnapshot WithoutFilters() => sheet.WithData(sheet.Data with { Filter = null, Tables = sheet.Data.Tables.Select(t => t.WithFilter(null)).ToImmutableArray() });
        if (WorksheetOutline.HasGroups(sheet) && sheet.Data.AllFilters.Any(f => f.Columns.Count > 0))
        {
            package.Loss("分组与活动筛选共存时，Excel 折叠/展开会改变 SUBTOTAL 的隐藏来源；当前保留原条件和源行状态，不重新解释此组合。",
                "worksheet/autoFilter-outline", path, range: string.Join(' ', sheet.Data.AllFilters.Select(f => f.Range.ToString())));
            return WithoutFilters();
        }
        ImmutableHashSet<int> filtered;
        try { filtered = WorksheetFilterEngine.Evaluate(sheet, styles, system, 0, allowSourceCaches: true); }
        catch (InvalidOperationException ex)
        {
            package.Loss(ex.Message + " 筛选隐藏来源无法确认，保留源行可见性。", "worksheet/autoFilter", path, range: string.Join(' ', sheet.Data.AllFilters.Select(f => f.Range.ToString())));
            return WithoutFilters();
        }
        var recorded = Rows((string?)metadata?.Attribute("filtered")); var manual = Rows((string?)metadata?.Attribute("manual")); var generated = Rows((string?)metadata?.Attribute("generated"));
        var trusted = metadata is not null && recorded.SetEquals(filtered) && filtered.All(r => sheet.Rows.GetValueOrDefault(r)?.Hidden == true) &&
            generated.IsSubsetOf(filtered) && !generated.Overlaps(manual) &&
            generated.All(r => sheet.Rows.TryGetValue(r, out var row) && row with { Hidden = false } == new RowMetadata());
        var rows = sheet.Rows.ToBuilder();
        foreach (var row in filtered)
        {
            if (!rows.TryGetValue(row, out var old) || !old.Hidden)
            {
                package.Loss("来源筛选条件与保存的隐藏行不一致，保留原包和源行状态；请在 Excel 重新应用筛选。", "worksheet/autoFilter-stale", path, range: string.Join(' ', sheet.Data.AllFilters.Select(f => f.Range.ToString())));
                return WithoutFilters();
            }
            var next = old with { Hidden = trusted && manual.Contains(row) };
            if (trusted ? generated.Contains(row) : next == new RowMetadata()) rows.Remove(row); else rows[row] = next;
        }
        return new(sheet.Id, sheet.Name, sheet.Cells, rows, sheet.Columns, sheet.MergedCells, sheet.Images, sheet.ShowGridLines,
            sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations, sheet.Data);
    }

    public static XElement? WriteFilter(WorksheetAutoFilter? filter, XElement? sort, XElement styles)
    {
        if (filter is null) return null;
        var root = new XElement(Main + "autoFilter", new XAttribute("ref", filter.Range));
        foreach (var (index, value) in filter.Columns.OrderBy(p => p.Key))
        {
            var column = new XElement(Main + "filterColumn", new XAttribute("colId", index - filter.Range.Left));
            if (!value.ShowButton) column.Add(new XAttribute("hiddenButton", 1));
            if (value.Color is { } color)
            {
                var dxf = new XElement(Main + "dxf", new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", color.Argb is null ? "none" : "solid"),
                        color.Argb is null ? null : new XElement(Main + "fgColor", new XAttribute("rgb", color.Argb)))));
                column.Add(new XElement(Main + "colorFilter", new XAttribute("dxfId", AddDifferentialStyle(styles, dxf)), new XAttribute("cellColor", color.Font ? 0 : 1)));
            }
            else if (value.Values is { } list) column.Add(new XElement(Main + "filters", value.IncludeBlank ? new XAttribute("blank", 1) : null,
                list.Select(v => new XElement(Main + "filter", new XAttribute("val", Encode(v))))));
            else if (value.First is not null)
                column.Add(new XElement(Main + "customFilters", value.UseAnd ? new XAttribute("and", 1) : null,
                    new[] { value.First, value.Second }.OfType<AutoFilterCondition>().Select(c => new XElement(Main + "customFilter",
                        new XAttribute("operator", Operators[(int)c.Comparison]), new XAttribute("val", Encode(c.Operand))))));
            root.Add(column);
        }
        root.Add(sort); return root;
    }

    private static int AddDifferentialStyle(XElement styles, XElement dxf)
    {
        var dxfs = styles.Element(Main + "dxfs");
        if (dxfs is null)
        {
            dxfs = new(Main + "dxfs");
            var after = styles.Elements().FirstOrDefault(e => e.Name.LocalName is "tableStyles" or "colors" or "extLst");
            if (after is null) styles.Add(dxfs); else after.AddBeforeSelf(dxfs);
        }
        var existing = dxfs.Elements().ToArray(); var id = Array.FindIndex(existing, e => XNode.DeepEquals(e, dxf));
        if (id < 0) { id = existing.Length; if (id >= 4096) throw new InvalidOperationException("差异样式数量超过 4096。"); dxfs.Add(dxf); }
        dxfs.SetAttributeValue("count", dxfs.Elements().Count()); return id;
    }

    public static XElement? WriteSort(WorksheetSortState? sort, XElement styles)
    {
        if (sort is null) return null;
        var root = new XElement(Main + "sortState", new XAttribute("ref", sort.Range), new XAttribute("caseSensitive", sort.Keys[0].CaseSensitive ? 1 : 0), new XAttribute("sortMethod", sort.SortMethod));
        foreach (var key in sort.Keys)
        {
            var item = new XElement(Main + "sortCondition", new XAttribute("ref", new CellRectangle(sort.Range.Top, key.Column, sort.Range.Bottom, key.Column)), new XAttribute("descending", key.Descending ? 1 : 0));
            if (!key.CustomOrder.IsEmpty) item.Add(new XAttribute("customList", Encode(string.Join(',', key.CustomOrder))));
            if (key.Basis != SortBasis.Value)
            {
                var dxf = new XElement(Main + "dxf", key.Basis == SortBasis.FontColor ? new XElement(Main + "font", new XElement(Main + "color", new XAttribute("rgb", key.PreferredColor!.ToUpperInvariant()))) :
                    new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", "solid"), new XElement(Main + "fgColor", new XAttribute("rgb", key.PreferredColor!.ToUpperInvariant())))));
                var dxfs = styles.Element(Main + "dxfs");
                if (dxfs is null)
                {
                    dxfs = new(Main + "dxfs");
                    var after = styles.Elements().FirstOrDefault(e => e.Name.LocalName is "tableStyles" or "colors" or "extLst");
                    if (after is null) styles.Add(dxfs); else after.AddBeforeSelf(dxfs);
                }
                var existing = dxfs.Elements().ToArray(); var id = Array.FindIndex(existing, e => XNode.DeepEquals(e, dxf));
                if (id < 0) { id = existing.Length; if (id >= 4096) throw new InvalidOperationException("差异样式数量超过 4096。"); dxfs.Add(dxf); }
                dxfs.SetAttributeValue("count", dxfs.Elements().Count());
                item.Add(new XAttribute("sortBy", key.Basis == SortBasis.FontColor ? "fontColor" : "cellColor"), new XAttribute("dxfId", id));
            }
            root.Add(item);
        }
        return root;
    }

    public static XElement? WriteHidden(WorksheetSnapshot sheet, ImmutableHashSet<int> filtered)
    {
        if (!sheet.Data.AllFilters.Any()) return null;
        string EncodeRows(IEnumerable<int> rows) => string.Join(',', rows.Order().Select(r => r + 1));
        return new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", Extension), new XElement(Quick + "rows",
            new XAttribute("filtered", EncodeRows(filtered)), new XAttribute("manual", EncodeRows(sheet.Rows.Where(p => p.Value.Hidden).Select(p => p.Key))),
            new XAttribute("generated", EncodeRows(filtered.Except(sheet.Rows.Keys))))));
    }

    public static XElement? WriteName(WorksheetSnapshot sheet, int index) => sheet.Data.Filter is not { } filter ? null : new(Main + "definedName",
        new XAttribute("name", "_xlnm._FilterDatabase"), new XAttribute("localSheetId", index), new XAttribute("hidden", 1),
        QuickExcel.Formulas.FormulaSyntax.Print(new QuickExcel.Formulas.ReferenceNode(sheet.Name, filter.Range.Top, filter.Range.Left, filter.Range.Bottom, filter.Range.Right,
            RowAbsolute: true, ColumnAbsolute: true, EndRowAbsolute: true, EndColumnAbsolute: true, IsRange: true)));
}
