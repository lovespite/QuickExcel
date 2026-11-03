using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>One visible row of a compatibility group; identical diagnostics are merged and located together.</summary>
public sealed record CompatibilityIssueView(
    string Feature,
    string Location,
    string Detail,
    string Impact,
    int Count,
    string Code);

public sealed record CompatibilityIssueGroup(
    CompatibilityKind Kind,
    string Label,
    string Summary,
    bool ExpandedByDefault,
    IReadOnlyList<CompatibilityIssueView> Items);

/// <summary>WPF-free description of the compatibility dialog and save prompt.</summary>
public sealed record CompatibilityReportView(
    string Title,
    string Headline,
    string Summary,
    IReadOnlyList<CompatibilityIssueGroup> Groups,
    IReadOnlyList<string> Notices)
{
    public int TotalCount => Groups.Sum(group => group.Items.Sum(item => item.Count));
    public int LossCount => Groups.Where(group => group.Kind == CompatibilityKind.PotentialLoss)
        .Sum(group => group.Items.Sum(item => item.Count));
    public bool HasLosses => LossCount > 0;
}

/// <summary>
/// Turns a capability report into impact-first copy for the report dialog and the save-time loss prompt.
/// This type stays free of WPF so the wording and grouping can be covered by Core tests.
/// </summary>
public static class CompatibilityReportPresentation
{
    private static readonly CompatibilityKind[] Order =
        [CompatibilityKind.PotentialLoss, CompatibilityKind.Rejected, CompatibilityKind.CannotCalculate,
         CompatibilityKind.PreservedOnly, CompatibilityKind.Editable];

    public static CompatibilityReportView Create(WorkbookCompatibilityReport report, bool saveMode,
        IEnumerable<string>? notices = null, bool canOverwrite = false)
    {
        // Inspection already classifies actual save losses. Preserved rich text,
        // links and notes must not become losses just because this is a save dialog.
        var groups = report.Diagnostics.GroupBy(diagnostic => diagnostic.Kind)
            .OrderBy(group => Array.IndexOf(Order, group.Key))
            .Select(group => BuildGroup(group.Key, group, saveMode, canOverwrite))
            .Where(group => group.Items.Count > 0)
            .ToArray();
        var losses = groups.Where(group => group.Kind == CompatibilityKind.PotentialLoss)
            .Sum(group => group.Items.Sum(item => item.Count));
        var title = saveMode ? "保存兼容性报告" : "兼容性报告";
        var headline = saveMode
            ? Count(losses, "处内容无法在编辑后保留")
            : report.OriginalPackageUnchanged
                ? "工作簿与导入时一致，可原样保存"
                : losses > 0 ? "当前编辑后，" + Count(losses, "处内容无法写入普通保存")
                : "未发现会阻止保存的格式损失";
        var summary = saveMode
            ? canOverwrite
                ? "另存受支持副本会保留原文件；覆盖会直接替换当前文件。两种保存都不包含下方“潜在损失”中的内容；取消则保留当前未保存的修改。"
                : "原文件不会被修改。继续会另存为受支持副本，副本不包含下方“潜在损失”中的内容；取消则保留当前未保存的修改。"
            : report.OriginalPackageUnchanged
                ? "原 XLSX 包可原样保留；普通保存不会重建未支持的部件。"
                : losses > 0
                    ? "保存时需明确选择如何处理损失；另存受支持副本可保留原文件中的内容。"
                    : "保存时仍会执行完整校验；下方同时列出仅保留、不可计算和已支持的内容。";
        var extra = (notices ?? []).Where(notice => !string.IsNullOrWhiteSpace(notice)).ToArray();
        return new(title, headline, summary, groups, extra);
    }

    public static string Impact(CompatibilityDiagnostic diagnostic, bool saveMode, bool canOverwrite = false)
    {
        if (diagnostic.Code.StartsWith("report.", StringComparison.Ordinal))
            return "这是报告容量提示：未逐项列出的内容仍受相同限制，保存保护覆盖全部内容。";
        if (diagnostic.Code == "workbook.formula-read-only")
            return "整个工作簿以只读的来源缓存打开，不能编辑、重算或做结构变换；原包与来源缓存保持完整。";
        if (diagnostic.Code == "formula.cache-summary")
            return "公式源码与缓存会保存；缺少当前有效结果的公式仍显示已有缓存，Excel 重算后结果可能不同。";
        if (diagnostic.Code is "text.rich-preserved" or "text.display-partial" or "text.attachments-preserved")
            return "数据可保留并保存；显示或编辑能力的限制见上述说明。";
        var specific = FeatureImpact(diagnostic.Feature);
        return diagnostic.Kind switch
        {
            CompatibilityKind.PotentialLoss => (specific ?? "该内容不属于已建模子集") +
                (saveMode ? canOverwrite
                    ? "；不会写入受支持副本，选择覆盖也会移除原文件中的该内容。"
                    : "；不会写入受支持副本，原文件保持不变。"
                    : "；修改后普通保存会被阻止，须明确另存副本或选择覆盖。"),
            CompatibilityKind.Rejected =>
                "该位置未通过导入检查，文件已阻止打开；请先在 Excel 中修复或删除该内容。",
            CompatibilityKind.CannotCalculate => (specific ?? "该公式不能在本机重新计算") +
                "；显示和保存的是已有缓存，Excel 重算后结果可能不同。",
            CompatibilityKind.PreservedOnly => (specific ?? "该内容只在原文件中保留") +
                "；只能在原文件中保留，另存编辑后的副本时不会按原样写入。",
            CompatibilityKind.Editable =>
                "可正常编辑并保存；仅代表已建模子集，不保证与 Excel 完全一致。",
            _ => "请查看原始说明。"
        };
    }

    /// <summary>Full, unabridged text for copy/support. Keeps the established report format.</summary>
    public static string PlainText(WorkbookCompatibilityReport report, IEnumerable<string>? notices = null)
    {
        var text = CompatibilityReportText.Format(report);
        var extra = (notices ?? []).Where(notice => !string.IsNullOrWhiteSpace(notice)).ToArray();
        return extra.Length == 0 ? text : text + "\n\n【其他提示】\n" + string.Join("\n", extra);
    }

    private static CompatibilityIssueGroup BuildGroup(CompatibilityKind kind,
        IEnumerable<CompatibilityDiagnostic> diagnostics, bool saveMode, bool canOverwrite)
    {
        var items = diagnostics
            .GroupBy(diagnostic => (diagnostic.Code, diagnostic.Feature, diagnostic.Message))
            .Select(group =>
            {
                var first = group.First();
                var locations = group.Select(Location).Where(location => location.Length > 0)
                    .Distinct(StringComparer.Ordinal).ToArray();
                return new CompatibilityIssueView(first.Feature, Locations(locations), first.Message,
                    Impact(first, saveMode, canOverwrite), group.Count(), first.Code);
            })
            .OrderByDescending(item => item.Count)
            .ThenBy(item => item.Feature, StringComparer.CurrentCulture)
            .ToArray();
        return new(kind, CompatibilityReportText.Label(kind), Summary(kind, saveMode, canOverwrite),
            kind is CompatibilityKind.PotentialLoss or CompatibilityKind.Rejected, items);
    }

    private static string Summary(CompatibilityKind kind, bool saveMode, bool canOverwrite) => kind switch
    {
        CompatibilityKind.PotentialLoss => saveMode
            ? canOverwrite ? "以下内容不会写入副本；选择覆盖时，也会从原文件中移除。" : "以下内容不会写入副本。原文件保持不变。"
            : "以下内容在修改后无法可靠保留，普通保存会被阻止。",
        CompatibilityKind.Rejected => "以下位置未通过导入检查，相关文件无法打开。",
        CompatibilityKind.CannotCalculate => "公式源码会保留，但本机不重算，显示与保存的是已有缓存。",
        CompatibilityKind.PreservedOnly => "以下内容仅支持保留或有限编辑；具体保存行为见各项说明。",
        CompatibilityKind.Editable => "已建模、可编辑并保存的子集。",
        _ => ""
    };

    private static string Location(CompatibilityDiagnostic diagnostic) => string.Join(" · ",
        new[] { diagnostic.SheetName, diagnostic.Range, diagnostic.Part }.Where(part => !string.IsNullOrWhiteSpace(part)));

    private static string Locations(IReadOnlyList<string> locations) => locations.Count switch
    {
        0 => "工作簿",
        <= 3 => string.Join("；", locations),
        _ => string.Join("；", locations.Take(2)) + $"；等 {locations.Count} 处"
    };

    private static string Count(int count, string tail) => count == 1 ? "有 1 " + tail : $"有 {count} {tail}";

    /// <summary>Short, user-facing consequence for the feature family. Never states a capability the report does not back.</summary>
    private static string? FeatureImpact(string feature)
    {
        var value = feature.ToLowerInvariant();
        if (value.StartsWith("part/"))
            return "未建模部件（如图表、透视表、宏或嵌入对象）没有可重建的模型";
        if (value.StartsWith("relationship/"))
            return "部件间的连接关系不会被重建，相关引用可能断开";
        if (value.StartsWith("formula/"))
        {
            if (value.Contains("version-prefix")) return "该公式使用较新的 Excel 函数前缀，本机不能计算";
            if (value.Contains("syntax")) return "该公式语法或名称引用不在当前支持范围内";
            if (value.Contains("dependency")) return "该公式依赖只读或暂不可用的结果";
            if (value.Contains("readonly")) return "该公式区域只读，不能编辑、重算或做结构变换";
            return "该公式不能在本机重新计算";
        }
        if (value.StartsWith("text."))
            return "文字仍可查看，但分段格式、链接或批注的编辑能力受限";
        if (value.Contains("pagesetup") || value.Contains("print") || value.Contains("headerfooter"))
            return "分页、打印区域或页眉页脚设置可能丢失";
        if (value.Contains("protection"))
            return "工作表或单元格保护可能不再生效";
        if (value.Contains("pane") || value.Contains("sheetview") || value.Contains("selection") ||
            value.Contains("frozen") || value.Contains("view"))
            return "冻结窗格、活动单元格或视图设置可能与原文件不同";
        if (value.Contains("font/") || value.Contains("fill/") || value.Contains("border/") ||
            value.Contains("cellxfs") || value.Contains("numfmt") || value.Contains("style") || value.Contains("alignment"))
            return "字体、填充、边框或数字格式可能回退为相近样式";
        if (value.Contains("conditionalformat") || value.Contains("datavalidation") || value.Contains("table") ||
            value.Contains("filter") || value.Contains("sort"))
            return "相关数据功能在副本中可能变为只读或失效";
        if (value.Contains("drawing") || value.Contains("image") || value.Contains("chart") ||
            value.Contains("shape") || value.Contains("picture"))
            return "图片、图形或图表可能不显示或位置改变";
        if (value.StartsWith("worksheet."))
            return "相关工作表设置可能只能保留在原文件中";
        if (value.Contains("shared"))
            return "共享内容在其他位置的使用可能一并受影响";
        return null;
    }
}
