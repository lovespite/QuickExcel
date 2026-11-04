using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    private static readonly XNamespace Chart15 = "http://schemas.microsoft.com/office/drawing/2012/chart";
    private const string FilterExtension = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
    private sealed record SeriesFilter(bool Hidden, ImmutableArray<int> Points);
    private static bool IsFilterExtension(XElement e) => e.Name == Chart + "ext" && string.Equals((string?)e.Attribute("uri"), FilterExtension, StringComparison.OrdinalIgnoreCase);
    private static void NoAttributes(XElement e)
    {
        if (e.Attributes().Any(a => !a.IsNamespaceDeclaration)) throw new NotSupportedException("未知图表筛选属性。");
    }
    private static string FilterText(XElement e)
    {
        NoAttributes(e);
        if (e.Elements().Count() != 1 || e.Element(Chart15 + "sqref") is not { HasElements: false } text) throw new NotSupportedException("未知图表筛选引用结构。");
        NoAttributes(text); return text.Value;
    }
    private static void RemoveEmptyExtensions(XElement root)
    {
        root.Descendants(Chart + "ext").Where(e => IsFilterExtension(e) && !e.HasElements).Remove();
        root.Descendants(Chart + "extLst").Where(e => !e.HasElements).Remove();
    }
    private static IEnumerable<string> ReferenceParts(string formula)
    {
        var text = formula.Trim().TrimStart('=');
        if (text.StartsWith('(') && text.EndsWith(')')) text = text[1..^1];
        if (text.Length == 0) yield break;
        var quoted = false; var start = 0;
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == '\'') { if (quoted && i + 1 < text.Length && text[i + 1] == '\'') { i++; continue; } quoted = !quoted; }
            if (text[i] != ',' || quoted) continue;
            yield return text[start..i]; start = i + 1;
        }
        if (quoted) throw new NotSupportedException("图表筛选引用引号不匹配。");
        yield return text[start..];
    }
    private static int[] VisibleIndices(string full, string visible)
    {
        var source = WorksheetChart.ParseReference(full); var count = WorksheetChart.Count(source);
        if (count > 4096) throw new NotSupportedException("图表筛选超过 4096 个点。");
        var result = new List<int>(); var previous = -1;
        foreach (var part in ReferenceParts(visible))
        {
            var r = WorksheetChart.ParseReference(part);
            if (!string.Equals(r.Sheet, source.Sheet, StringComparison.OrdinalIgnoreCase) || r.Row < source.Row || r.EndRow > source.EndRow || r.Column < source.Column || r.EndColumn > source.EndColumn)
                throw new NotSupportedException("图表筛选引用超出完整源区域。");
            var first = source.EndRow > source.Row ? r.Row - source.Row : r.Column - source.Column;
            for (var i = first; i < first + WorksheetChart.Count(r); i++)
            { if (i <= previous || i >= count) throw new NotSupportedException("图表筛选引用重复或顺序无效。"); result.Add(i); previous = i; }
        }
        return result.ToArray();
    }
    private static string PointReference(string formula, int first, int last)
    {
        var r = WorksheetChart.ParseReference(formula); var vertical = r.EndRow > r.Row;
        return FormulaSyntax.Print(r with { Row = r.Row + (vertical ? first : 0), EndRow = r.Row + (vertical ? last : 0),
            Column = r.Column + (vertical ? 0 : first), EndColumn = r.Column + (vertical ? 0 : last),
            RowAbsolute = true, ColumnAbsolute = true, EndRowAbsolute = true, EndColumnAbsolute = true, IsRange = first != last });
    }
    private static string VisibleReference(string formula, IReadOnlySet<int> hidden)
    {
        var count = WorksheetChart.Count(WorksheetChart.ParseReference(formula)); var parts = new List<string>();
        for (var i = 0; i < count; i++)
        {
            if (hidden.Contains(i)) continue;
            var first = i; while (i + 1 < count && !hidden.Contains(i + 1)) i++;
            parts.Add(PointReference(formula, first, i).TrimStart('='));
        }
        return parts.Count > 1 ? "(" + string.Join(",", parts) + ")" : parts.FirstOrDefault() ?? "";
    }
    // Normalize supported Office filtering into full references before the ordinary closed-schema reader runs.
    // Hidden series and hidden point styles are validated by exactly the same reader as visible content.
    private static Dictionary<XElement, SeriesFilter> NormalizeStandardFilters(XElement root)
    {
        var result = new Dictionary<XElement, SeriesFilter>();
        // Office emits explicit false/default flags even for ordinary flat pie points.
        // Non-default flags remain in place and are rejected by the closed-schema reader.
        root.Descendants().Where(e => (e.Parent?.Name == Chart + "dPt" && e.Name.Namespace == Chart || e.Parent?.Name == Chart15 + "categoryFilterException" && e.Name.Namespace == Chart15) &&
            e.Name.LocalName is "bubble3D" or "invertIfNegative" or "explosion" && !e.HasElements &&
            e.Attributes().Count(a => !a.IsNamespaceDeclaration) == 1 && (string?)e.Attribute("val") is "0" or "false").Remove();
        foreach (var ext in root.Descendants().Where(IsFilterExtension))
            if (!ext.HasElements || ext.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "uri")) throw new NotSupportedException("无效图表筛选扩展。");
        foreach (var graph in root.Element(Chart + "chart")?.Element(Chart + "plotArea")?.Elements().Where(e => e.Name.LocalName is "barChart" or "lineChart" or "pieChart" or "scatterChart").ToArray() ?? [])
        {
            var hiddenSeries = new HashSet<XElement>();
            var wrapperName = graph.Name.LocalName switch { "barChart" => "filteredBarSeries", "lineChart" => "filteredLineSeries", "pieChart" => "filteredPieSeries", _ => "filteredScatterSeries" };
            foreach (var ext in graph.Element(Chart + "extLst")?.Elements().Where(IsFilterExtension).ToArray() ?? [])
            foreach (var wrapper in ext.Elements().ToArray())
            {
                NoAttributes(wrapper);
                if (wrapper.Name != Chart15 + wrapperName || wrapper.Elements().Count() != 1 || wrapper.Element(Chart15 + "ser") is not { } ser)
                    throw new NotSupportedException("未知筛选图表系列。");
                NoAttributes(ser); ser.Remove(); ser.Name = Chart + "ser"; graph.Add(ser); hiddenSeries.Add(ser); wrapper.Remove();
            }
            var series = graph.Elements(Chart + "ser").OrderBy(s => int.Parse(Value(s, "order") ?? "0", CultureInfo.InvariantCulture)).ToArray();
            if (series.Length > 16 || series.Select(s => Value(s, "idx")).Distinct().Count() != series.Length || series.Select(s => Value(s, "order")).Distinct().Count() != series.Length)
                throw new NotSupportedException("图表系列序号无效。");
            foreach (var ser in series)
            {
                var maps = new Dictionary<string, int[]>();
                foreach (var container in ser.Elements().Where(e => e.Name.LocalName is "tx" or "cat" or "val" or "xVal" or "yVal"))
                foreach (var reference in container.Elements().Where(e => e.Name == Chart + "strRef" || e.Name == Chart + "numRef"))
                {
                    string? full = null, formula = null, level = null;
                    foreach (var ext in reference.Element(Chart + "extLst")?.Elements().Where(IsFilterExtension).ToArray() ?? [])
                    foreach (var field in ext.Elements().ToArray())
                    {
                        var text = FilterText(field);
                        if (field.Name == Chart15 + "fullRef" && full is null) full = text;
                        else if (field.Name == Chart15 + "formulaRef" && formula is null) formula = text;
                        else if (field.Name == Chart15 + "levelRef" && level is null) level = text;
                        else throw new NotSupportedException("未知或重复图表筛选引用。");
                        field.Remove();
                    }
                    var visible = reference.Element(Chart + "f")?.Value ?? formula;
                    if (visible is null) throw new NotSupportedException("图表筛选引用缺失。");
                    if (formula is not null && reference.Element(Chart + "f") is { } f && f.Value != formula) throw new NotSupportedException("图表筛选引用冲突。");
                    full ??= level ?? visible;
                    if (level is not null && !VisibleIndices(full, level).SequenceEqual(Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(full)))))
                        throw new NotSupportedException("多层图表分类尚未支持。");
                    var positions = VisibleIndices(full, visible); maps[container.Name.LocalName] = positions;
                    if (container.Name.LocalName == "tx" && positions.Length != 1) throw new NotSupportedException("被筛选的系列标题尚未支持。");
                    reference.Elements(Chart + "f").Remove(); reference.AddFirst(new XElement(Chart + "f", full.TrimStart('=')));
                }
                var catName = graph.Name.LocalName == "scatterChart" ? "xVal" : "cat"; var valueName = graph.Name.LocalName == "scatterChart" ? "yVal" : "val";
                if (!maps.TryGetValue(catName, out var cats) || !maps.TryGetValue(valueName, out var values) || !cats.SequenceEqual(values)) throw new NotSupportedException("图表分类和数值筛选必须对应。");
                var fullValues = ser.Element(Chart + valueName)!.Element(Chart + "numRef")!.Element(Chart + "f")!.Value;
                var count = WorksheetChart.Count(WorksheetChart.ParseReference(fullValues));
                foreach (var point in ser.Elements(Chart + "dPt"))
                {
                    var index = int.Parse(Value(point, "idx") ?? "-1", CultureInfo.InvariantCulture);
                    if (index < 0 || index >= values.Length) throw new NotSupportedException("图表数据点筛选序号无效。");
                    point.Element(Chart + "idx")!.SetAttributeValue("val", values[index]);
                }
                foreach (var ext in ser.Element(Chart + "extLst")?.Elements().Where(IsFilterExtension).ToArray() ?? [])
                foreach (var exceptions in ext.Elements().ToArray())
                {
                    NoAttributes(exceptions);
                    if (exceptions.Name != Chart15 + "categoryFilterExceptions") throw new NotSupportedException("未知筛选数据点结构。");
                    foreach (var exception in exceptions.Elements())
                    {
                        NoAttributes(exception);
                        if (exception.Name != Chart15 + "categoryFilterException" || exception.Elements().Any(e => e.Name != Chart15 + "sqref" && e.Name != Chart15 + "spPr") ||
                            exception.Elements(Chart15 + "sqref").Count() != 1 || exception.Elements(Chart15 + "spPr").Count() > 1) throw new NotSupportedException("未建模筛选数据点格式。");
                        var sqref = exception.Element(Chart15 + "sqref")!; NoAttributes(sqref); if (sqref.HasElements) throw new NotSupportedException("无效筛选数据点引用。");
                        var indices = VisibleIndices(fullValues, sqref.Value);
                        if (indices.Length != 1 || values.Contains(indices[0]) || ser.Elements(Chart + "dPt").Any(p => Value(p, "idx") == indices[0].ToString(CultureInfo.InvariantCulture))) throw new NotSupportedException("筛选数据点引用冲突。");
                        if (exception.Element(Chart15 + "spPr") is { } shape)
                        {
                            var copy = new XElement(shape) { Name = Chart + "spPr" }; var point = new XElement(Chart + "dPt", Val("idx", indices[0]), copy);
                            var before = ser.Elements().FirstOrDefault(e => e.Name.LocalName is "dLbls" or "cat" or "val" or "xVal" or "yVal" or "smooth" or "extLst");
                            if (before is null) ser.Add(point); else before.AddBeforeSelf(point);
                        }
                    }
                    exceptions.Remove();
                }
                result[ser] = new(hiddenSeries.Contains(ser), [.. Enumerable.Range(0, count).Except(values)]);
                ser.Remove(); graph.Add(ser);
            }
        }
        RemoveEmptyExtensions(root);
        return result;
    }
    private static XElement FilterField(string name, string text) => new(Chart15 + name, new XElement(Chart15 + "sqref", text.TrimStart('=')));
    private static void AddFilterExtension(XElement parent, params XElement[] fields)
    {
        if (fields.Length == 0) return;
        var list = parent.Element(Chart + "extLst"); if (list is null) parent.Add(list = new XElement(Chart + "extLst"));
        var ext = list.Elements().FirstOrDefault(IsFilterExtension); if (ext is null) list.Add(ext = new XElement(Chart + "ext", new XAttribute("uri", FilterExtension)));
        ext.Add(fields);
    }
    private static void WriteStandardFilters(XElement graph, WorksheetChart model)
    {
        var nodes = graph.Elements(Chart + "ser").ToArray();
        for (var i = 0; i < nodes.Length; i++)
        {
            var ser = nodes[i]; var s = model.Series[i]; var hidden = s.HiddenPoints.ToHashSet();
            var selected = Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Where(p => !hidden.Contains(p)).ToArray();
            foreach (var reference in ser.Elements().Where(e => e.Name.LocalName is "tx" or "cat" or "val" or "xVal" or "yVal").SelectMany(e => e.Elements()).Where(e => e.Name == Chart + "strRef" || e.Name == Chart + "numRef"))
            {
                var f = reference.Element(Chart + "f")!; var full = f.Value; var isName = reference.Parent!.Name == Chart + "tx";
                var visible = isName ? full : VisibleReference(full, hidden);
                var fields = new List<XElement>(); if (!isName && hidden.Count > 0) fields.Add(FilterField("fullRef", full));
                if (s.IsFiltered) { fields.Add(FilterField("formulaRef", visible)); f.Remove(); } else f.Value = visible;
                AddFilterExtension(reference, fields.ToArray());
                if (!isName && hidden.Count > 0)
                {
                    var cache = reference.Elements().First(e => e.Name.LocalName is "strCache" or "numCache");
                    foreach (var point in cache.Elements(Chart + "pt").ToArray())
                    {
                        var original = int.Parse((string)point.Attribute("idx")!, CultureInfo.InvariantCulture); var position = Array.IndexOf(selected, original);
                        if (position < 0) point.Remove(); else point.SetAttributeValue("idx", position);
                    }
                    cache.Element(Chart + "ptCount")!.SetAttributeValue("val", selected.Length);
                }
            }
            var exceptions = new List<XElement>();
            foreach (var point in ser.Elements(Chart + "dPt").ToArray())
            {
                var original = int.Parse(Value(point, "idx")!, CultureInfo.InvariantCulture); var position = Array.IndexOf(selected, original);
                if (position >= 0) point.Element(Chart + "idx")!.SetAttributeValue("val", position);
                else
                {
                    if (point.Element(Chart + "spPr") is { } shape) exceptions.Add(new XElement(Chart15 + "categoryFilterException", new XElement(Chart15 + "sqref", PointReference(s.Values, original, original).TrimStart('=')), new XElement(shape) { Name = Chart15 + "spPr" },
                        model.Kind == WorksheetChartKind.Column ? new XElement(Chart15 + "invertIfNegative", new XAttribute("val", 0)) : null));
                    point.Remove();
                }
            }
            if (exceptions.Count > 0) AddFilterExtension(ser, new XElement(Chart15 + "categoryFilterExceptions", exceptions));
            if (!s.IsFiltered) continue;
            ser.Remove(); ser.Name = Chart15 + "ser";
            var wrapper = model.Kind switch { WorksheetChartKind.Column => "filteredBarSeries", WorksheetChartKind.Pie => "filteredPieSeries", WorksheetChartKind.Scatter => "filteredScatterSeries", _ => "filteredLineSeries" };
            AddFilterExtension(graph, new XElement(Chart15 + wrapper, ser));
        }
    }
}
