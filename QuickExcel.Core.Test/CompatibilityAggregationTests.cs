using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class CompatibilityAggregationTests
{
    [Fact]
    public void ReportCountsSemanticTypesAndUniqueRowsInsteadOfMessagesOrCellLocations()
    {
        CompatibilityDiagnostic Cell(string feature, string sheet, string range, string message, string code = "xlsx.unmodeled") =>
            new(code, CompatibilityKind.PotentialLoss, feature, message, "xl/styles.xml", sheet, range);
        var report = new WorkbookCompatibilityReport(false, [
            new("xlsx.unmodeled", CompatibilityKind.PotentialLoss, "font/outline", "source", "xl/styles.xml", Range: "cellXfs[0]"),
            Cell("font/outline · cellXfs[0]", "第一表", "A1", "first message", "xlsx.affected-location"),
            Cell("font/outline · cellXfs[1]", "第一表", "B1", "different message", "xlsx.affected-location"),
            Cell("font/outline", "第一表", "row:2", "row-specific message"),
            Cell("font/outline", "第二表", "A1", "other sheet"),
            Cell("alignment/@unknown", "第一表", "C2", "second feature on an already counted row"),
            new("report.locations-truncated", CompatibilityKind.PotentialLoss, "报告容量", "samples truncated")]);

        var view = CompatibilityReportPresentation.Create(report, true);
        var group = Assert.Single(view.Groups);
        Assert.Equal(2, view.TotalCount); Assert.Equal(2, view.LossCount);
        Assert.Equal(3, view.AffectedRowCount); Assert.Equal(3, group.AffectedRowCount);
        Assert.Equal(3, Assert.Single(group.Items, item => item.Feature == "font/outline").AffectedRowCount);
        Assert.Equal(1, Assert.Single(group.Items, item => item.Feature == "alignment/@unknown").AffectedRowCount);
        Assert.Equal(new[] { "samples truncated" }, view.Notices);
        Assert.Contains("2 类内容无法在编辑后保留", view.Headline);
        var text = CompatibilityReportText.Format(report);
        Assert.Contains("【潜在损失】2 类", text); Assert.Contains("受影响 3 行", text);
        Assert.Contains("【报告提示】", text);
    }

    [Fact]
    public void BoundedSamplesKeepLaterTypesAndExactRowsAcrossMultipleCellsAndSheets()
    {
        var collector = new CompatibilityDiagnosticCollector();
        for (var row = 1; row <= 1500; row++)
        {
            collector.Add(new("xlsx.unmodeled", CompatibilityKind.PreservedOnly, "row/@unknown", "row " + row,
                "xl/worksheets/sheet1.xml", "数据", "A" + row));
            collector.Add(new("xlsx.unmodeled", CompatibilityKind.PreservedOnly, "row/@unknown", "same row other cell " + row,
                "xl/worksheets/sheet1.xml", "数据", "B" + row));
        }
        collector.Add(new("xlsx.unmodeled", CompatibilityKind.PreservedOnly, "row/@unknown", "second sheet",
            "xl/worksheets/sheet2.xml", "第二表", "A1"));
        collector.Add(new("xlsx.unmodeled", CompatibilityKind.PreservedOnly, "cell/@late", "new type after sample cap",
            "xl/worksheets/sheet1.xml", "数据", "B1495:B1500"));
        var diagnostics = collector.Diagnostics;
        Assert.Equal(1024, diagnostics.Count(d => d.Range is not null));
        Assert.Contains(diagnostics, d => d.Code == "report.locations-truncated");
        var report = new WorkbookCompatibilityReport(true, diagnostics);
        var group = Assert.Single(CompatibilityReportPresentation.Create(report, false).Groups);
        Assert.Equal(2, group.Items.Count); Assert.Equal(1501, group.AffectedRowCount);
        Assert.Equal(1501, Assert.Single(group.Items, item => item.Feature == "row/@unknown").AffectedRowCount);
        Assert.Equal(6, Assert.Single(group.Items, item => item.Feature == "cell/@late").AffectedRowCount);
        Assert.All(group.Items, item => Assert.True(item.Coverage.LocationsTruncated));
        Assert.Equal(diagnostics, collector.Diagnostics);
    }

    [Fact]
    public void OverlappingRangesAndDuplicateRowsAreMergedWithoutExpandingColumnScopes()
    {
        var collector = new CompatibilityDiagnosticCollector();
        foreach (var range in new[] { "A10:C20", "$B$3:$D$12", "row:7", "A1:B2", "columns:1:3" })
            collector.Add(new("test", CompatibilityKind.PreservedOnly, "feature", range, SheetName: "数据", Range: range));
        collector.Add(new("test", CompatibilityKind.PreservedOnly, "feature", "shared source", "xl/styles.xml", Range: "cellXfs[0]"));
        var item = Assert.Single(Assert.Single(CompatibilityReportPresentation.Create(new(true, collector.Diagnostics), false).Groups).Items);
        Assert.Equal(20, item.AffectedRowCount); Assert.True(item.Coverage.HasNonRowLocations);
        Assert.Equal(new CompatibilityRowSpan("数据", 1, 20), Assert.Single(item.Coverage.Rows.Spans));
    }
}
