using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

internal static partial class XlsxCharts
{
    private static readonly XNamespace QuickChart = "urn:quickexcel:chart-filters:1";
    private const string ExtendedFilterExtension = "{C660D6BA-5394-43F8-9175-214870530CF8}";
    private static bool IsExtendedFilter(XElement e) => e.Name == ChartEx + "ext" && (string?)e.Attribute("uri") == ExtendedFilterExtension;
    private static XElement ExtendedSeries(XElement root) => root.Element(ChartEx + "chart")!.Element(ChartEx + "plotArea")!.Element(ChartEx + "plotAreaRegion")!.Element(ChartEx + "series")!;
    private static void ProjectExtendedSeries(XElement series, int[] selected, bool hidden)
    {
        series.SetAttributeValue("hidden", hidden ? "1" : null);
        foreach (var point in series.Elements(ChartEx + "dataPt").Concat(series.Element(ChartEx + "layoutPr")?.Element(ChartEx + "subtotals")?.Elements() ?? []).ToArray())
        {
            var attribute = point.Name.LocalName == "idx" ? "val" : "idx";
            var original = int.Parse((string)point.Attribute(attribute)!, CultureInfo.InvariantCulture); var index = Array.IndexOf(selected, original);
            if (index < 0) point.Remove(); else point.SetAttributeValue(attribute, index);
        }
        if (series.Element(ChartEx + "layoutPr")?.Element(ChartEx + "subtotals") is { HasElements: false } empty) empty.Remove();
    }
    private static XElement Canonical(XElement node) => new(node.Name, node.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString()).Select(a => new XAttribute(a)),
        node.HasElements ? node.Elements().Select(Canonical) : node.Value);
    private static SeriesFilter NormalizeExtendedFilters(XElement root)
    {
        var outerSeries = ExtendedSeries(root);
        var hidden = (string?)outerSeries.Attribute("hidden") is "1" or "true";
        var extensions = root.Element(ChartEx + "extLst")?.Elements().ToArray() ?? [];
        if (extensions.Length == 0) return new(hidden, []);
        if (extensions.Length != 1 || !IsExtendedFilter(extensions[0])) throw new NotSupportedException("未知 ChartEx 筛选扩展。");
        var ext = extensions[0];
        if (ext.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "uri") || ext.Elements().Count() != 1 || ext.Element(QuickChart + "filter") is not { } filter ||
            filter.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "hiddenPoints") || filter.Elements().Count() != 2 ||
            filter.Element(ChartEx + "chartData") is not { } fullData || filter.Element(ChartEx + "series") is not { } fullSeries)
            throw new NotSupportedException("无效 ChartEx 筛选扩展。");
        var points = ((string?)filter.Attribute("hiddenPoints") ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(s => int.Parse(s, CultureInfo.InvariantCulture)).ToImmutableArray();
        var dimensions = fullData.Element(ChartEx + "data")?.Elements().ToArray() ?? [];
        var visibleDimensions = root.Element(ChartEx + "chartData")?.Element(ChartEx + "data")?.Elements().ToArray() ?? [];
        if (dimensions.Length != 2 || visibleDimensions.Length != 2 || fullData.Elements().Count() != 1 || root.Element(ChartEx + "chartData")!.Elements().Count() != 1) throw new NotSupportedException("无效 ChartEx 筛选数据。");
        int[]? selected = null;
        foreach (var dimension in dimensions)
        {
            var full = dimension.Element(ChartEx + "f")?.Value ?? throw new NotSupportedException("ChartEx 完整引用缺失。");
            var visible = visibleDimensions.SingleOrDefault(d => d.Name == dimension.Name && (string?)d.Attribute("type") == (string?)dimension.Attribute("type"))?.Element(ChartEx + "f")?.Value ?? throw new NotSupportedException("ChartEx 筛选引用缺失。");
            var count = WorksheetChart.Count(WorksheetChart.ParseReference(full));
            if (points.Distinct().Count() != points.Length || points.Any(i => i < 0 || i >= count)) throw new NotSupportedException("ChartEx 筛选序号无效。");
            var indices = VisibleIndices(full, visible);
            if (!indices.SequenceEqual(Enumerable.Range(0, count).Except(points)) || selected is not null && !selected.SequenceEqual(indices)) throw new NotSupportedException("ChartEx 筛选引用冲突。");
            selected = indices;
        }
        var projected = new XElement(fullSeries); ProjectExtendedSeries(projected, selected!, hidden);
        if (!XNode.DeepEquals(Canonical(projected), Canonical(outerSeries))) throw new NotSupportedException("ChartEx 筛选样式与完整系列不一致。");
        root.Element(ChartEx + "chartData")!.ReplaceWith(new XElement(fullData)); outerSeries.ReplaceWith(new XElement(fullSeries)); root.Element(ChartEx + "extLst")!.Remove();
        // Validate the hidden styles as well as the visible projection.
        ValidateExtended(root);
        return new(hidden, points);
    }
    private static void WriteExtendedFilters(XElement root, WorksheetChart model)
    {
        var s = model.Series[0]; var series = ExtendedSeries(root);
        series.SetAttributeValue("hidden", s.IsFiltered ? "1" : null);
        if (s.HiddenPoints.IsEmpty) return;
        var hidden = s.HiddenPoints.ToHashSet(); var selected = Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Where(i => !hidden.Contains(i)).ToArray();
        var data = root.Element(ChartEx + "chartData")!; var fullData = new XElement(data); fullData.Descendants(ChartEx + "lvl").Remove();
        var fullSeries = new XElement(series); fullSeries.SetAttributeValue("hidden", null);
        foreach (var dimension in data.Element(ChartEx + "data")!.Elements())
        {
            var f = dimension.Element(ChartEx + "f")!; f.Value = VisibleReference(f.Value, hidden);
            foreach (var level in dimension.Elements(ChartEx + "lvl"))
            {
                foreach (var point in level.Elements(ChartEx + "pt").ToArray())
                {
                    var index = Array.IndexOf(selected, int.Parse((string)point.Attribute("idx")!, CultureInfo.InvariantCulture));
                    if (index < 0) point.Remove(); else point.SetAttributeValue("idx", index);
                }
                level.SetAttributeValue("ptCount", selected.Length);
            }
        }
        ProjectExtendedSeries(series, selected, s.IsFiltered);
        root.Add(new XElement(ChartEx + "extLst", new XElement(ChartEx + "ext", new XAttribute("uri", ExtendedFilterExtension),
            new XElement(QuickChart + "filter", new XAttribute("hiddenPoints", string.Join(" ", s.HiddenPoints.Order())), fullData, fullSeries))));
    }
}
