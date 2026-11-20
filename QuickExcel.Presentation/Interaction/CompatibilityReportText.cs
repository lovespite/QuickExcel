using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>Text shared by the actual dialog and focused, WPF-free presentation tests.</summary>
public static class CompatibilityReportText
{
    public static string Format(WorkbookCompatibilityReport report)
    {
        var lines = new List<string> {
            report.OriginalPackageUnchanged
                ? UiText.Get("工作簿内容与导入时一致，普通保存 XLSX/XLSQ 可保留原 XLSX 包；仅保留不代表可编辑写回。")
                : report.Diagnostics.Any(d => d.Kind == CompatibilityKind.PotentialLoss)
                    ? UiText.Get("当前内容未走原包不变保留路径。潜在损失会阻止普通保存，须明确另存受支持副本，或在保存报告允许时确认覆盖原文件。")
                    : UiText.Get("当前报告未列出格式损失阻止项；保存时仍会执行完整校验。"),
            UiText.Get("文件特性定位对应导入时的部件/工作表/范围；公式诊断对应当前单元格。内置不可计算与保存损失分别报告。")
        };
        foreach (var group in report.Diagnostics.Where(d => !d.IsReportNotice).GroupBy(d => d.Kind).OrderBy(g => g.Key))
        {
            var types = group.GroupBy(d => (d.TypeCode, d.TypeFeature, d.TypeNamespace)).ToArray();
            lines.Add(UiText.Format($"【{Label(group.Key)}】{types.Length} 类") + " · " +
                CompatibilityReportPresentation.RowSummary(CompatibilityDiagnosticCollector.Summarize(group)));
            foreach (var type in types)
            {
                lines.Add(type.Key.TypeFeature + " · " + CompatibilityReportPresentation.RowSummary(CompatibilityDiagnosticCollector.Summarize(type)));
                lines.AddRange(type.Select(d => string.Join(" · ", new[] { d.SheetName, d.Range, d.Part, d.Feature }
                    .Where(s => !string.IsNullOrEmpty(s))) + "：" + d.Message));
            }
        }
        var notices = report.Diagnostics.Where(d => d.IsReportNotice).Select(d => UiText.Get(d.Message)).Distinct(StringComparer.Ordinal).ToArray();
        if (notices.Length > 0) lines.Add(UiText.Get("【报告提示】") + "\n" + string.Join("\n", notices));
        return string.Join("\n\n", lines);
    }

    public static string Rejection(CompatibilityDiagnostic diagnostic) => UiText.Get("【拒绝】") + string.Join(" · ",
        new[] { diagnostic.SheetName, diagnostic.Range, diagnostic.Part, diagnostic.Feature }.Where(s => !string.IsNullOrEmpty(s))) + "：" + diagnostic.Message;

    public static string Label(CompatibilityKind kind) => kind switch
    {
        CompatibilityKind.Editable => UiText.Get("可编辑（已支持子集）"),
        CompatibilityKind.PreservedOnly => UiText.Get("仅保留"),
        CompatibilityKind.CannotCalculate => UiText.Get("内置不可计算／待验证"),
        CompatibilityKind.PotentialLoss => UiText.Get("潜在损失"),
        CompatibilityKind.Rejected => UiText.Get("拒绝"),
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };
}
