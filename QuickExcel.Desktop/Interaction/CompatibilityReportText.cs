using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>Text shared by the actual dialog and focused, WPF-free presentation tests.</summary>
public static class CompatibilityReportText
{
    public static string Format(WorkbookCompatibilityReport report)
    {
        var lines = new List<string> {
            report.OriginalPackageUnchanged
                ? "工作簿内容与导入时一致，普通保存 XLSQ 可保留原 XLSX 包；仅保留不代表可编辑写回。"
                : report.Diagnostics.Any(d => d.Kind == CompatibilityKind.PotentialLoss)
                    ? "当前内容未走原包不变保留路径。潜在损失会阻止普通保存，须明确另存受支持副本。"
                    : "当前报告未列出格式损失阻止项；保存时仍会执行完整校验。",
            "文件特性定位对应导入时的部件/工作表/范围；公式诊断对应当前单元格。内置不可计算与保存损失分别报告。"
        };
        foreach (var group in report.Diagnostics.GroupBy(d => d.Kind).OrderBy(g => g.Key))
        {
            lines.Add($"【{Label(group.Key)}】{group.Count()} 项");
            lines.AddRange(group.Select(d => string.Join(" · ", new[] { d.SheetName, d.Range, d.Part, d.Feature }
                .Where(s => !string.IsNullOrEmpty(s))) + "：" + d.Message));
        }
        return string.Join("\n\n", lines);
    }

    public static string Label(CompatibilityKind kind) => kind switch
    {
        CompatibilityKind.Editable => "可编辑（已支持子集）",
        CompatibilityKind.PreservedOnly => "仅保留",
        CompatibilityKind.CannotCalculate => "内置不可计算／待验证",
        CompatibilityKind.PotentialLoss => "潜在损失",
        CompatibilityKind.Rejected => "拒绝",
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };
}
