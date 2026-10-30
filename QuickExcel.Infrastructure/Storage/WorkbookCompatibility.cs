using QuickExcel.Domain;
using QuickExcel.Core;
using QuickExcel.Formulas;

namespace QuickExcel.Storage;

public sealed class WorkbookFeatureRejectedException(CompatibilityDiagnostic diagnostic, Exception? inner = null)
    : NotSupportedException($"拒绝导入：{diagnostic.SheetName}!{diagnostic.Range} · {diagnostic.Feature}：{diagnostic.Message}", inner)
{
    public CompatibilityDiagnostic Diagnostic { get; } = diagnostic;
}

public sealed partial class TypedWorkbookStorage
{
    /// <summary>Read-only inspection of the current snapshot. Does not calculate, load scripts or change save policy.</summary>
    public static WorkbookCompatibilityReport InspectCompatibility(WorkbookSnapshot snapshot, CancellationToken cancellationToken = default)
    {
        var inspected = PrepareReadOnlyFormulaSave(snapshot, false);
        var unchanged = inspected.FileState is { } state && ContentEquals(inspected, state.OriginalModel);
        var diagnostics = (snapshot.FileState?.CompatibilityDiagnostics ?? []).Select(d => d with {
            Kind = unchanged ? d.Kind : CompatibilityKind.PotentialLoss
        }).ToList();
        if (snapshot.FileState is { FormulaReadOnly: true } source)
        {
            diagnostics.Add(new("workbook.formula-read-only", CompatibilityKind.PreservedOnly, "只读公式工作簿", WorkbookFileState.FormulaReadOnlyMessage));
            foreach (var formula in source.PreservedFormulas.Take(1024))
                diagnostics.Add(new("formula.preserved-read-only", CompatibilityKind.PreservedOnly, "formula/" + formula.Kind,
                    "公式及结果范围依赖原包，仅作只读投影；源码、公式属性、相关部件和来源缓存整包保留。共享从属格显示主表达式，数组结果区域中的值也属于来源缓存。",
                    formula.Part, formula.SheetName, formula.Range ?? formula.Cell));
            if (source.PreservedFormulas.Length > 1024) diagnostics.Add(new("report.preserved-truncated", CompatibilityKind.PreservedOnly, "报告容量", "只展示前 1024 个只读公式位置；整簿只读限制仍覆盖全部公式。"));
        }
        else diagnostics.Add(new("workbook.modeled-subset", CompatibilityKind.Editable, "已建模内容",
            "单元格类型、普通公式源码和基础格式可在已支持子集内编辑；不支持的函数及依赖保留旧缓存，未知引用语义阻止结构变换。这不代表全部 Excel 功能已支持。"));
        var formulaCount = 0;
        var cachedCount = 0;
        var staleCount = 0;
        var formulaLocations = 0;
        var truncated = false;
        var richCount = 0; var richWarnings = 0;
        foreach (var sheet in snapshot.Sheets)
        foreach (var (position, cell) in sheet.Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.Content.RichText is { } rich)
            {
                richCount++;
                var style = snapshot.Styles[cell.Content.StyleId];
                if (richWarnings < 1024 && (style.ShrinkToFit || rich.Runs.Any(r => r.Font?.Vertical is "superscript" or "subscript" || r.Font?.Underline is "singleAccounting" or "doubleAccounting" ||
                        r.Font is { } font && (font.Name is null && font.Scheme is null || font.Size is null) &&
                        (style.FontName != snapshot.Styles[0].FontName || style.FontSize != snapshot.Styles[0].FontSize)) ||
                    CellDisplayFormatter.Format(cell.Value, style, snapshot.DateSystem) != rich.Text))
                {
                    richWarnings++;
                    diagnostics.Add(new("text.display-partial", CompatibilityKind.PreservedOnly, "富文本显示限制",
                        "分段数据保存保留；上下标/会计下划线、缺少字体属性的分段回退、混合字号缩小填充或改变文字的数字格式显示尚未完全对齐。", SheetName: sheet.Name, Range: TypedWorkbookXml.Reference(position)));
                }
            }
            if (cell.Content.Kind != CellContentKind.Formula) continue;
            formulaCount++;
            if (cell.Calculation?.Value is not null) cachedCount++;
            if (cell.Calculation is not { Value: not null } calculation || calculation.InputRevision != snapshot.Revision ||
                calculation.State is not (CalculationState.Fresh or CalculationState.Error)) staleCount++;
            if (formulaLocations >= 1024) { truncated = true; continue; }
            string? message = null;
            var code = "formula.unsupported-function";
            var feature = "formula/function";
            try
            {
                var expression = FormulaBinding.Bind(FormulaSyntax.Parse(cell.Content.FormulaText), snapshot, sheet.Id);
                var unknown = FormulaSyntax.Walk(expression).OfType<FunctionNode>().Select(f => f.Name)
                    .Where(name => !StandardFunctions.Signatures.ContainsKey(name)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                if (unknown.Length > 0)
                {
                    if (unknown.Any(n => n.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) || n.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase))) feature = "formula/version-prefix";
                    message = "内置计算器不支持函数：" + string.Join("、", unknown) +
                        "。公式源码仍保留；若为自定义脚本函数，须在显式启用后核对注册结果，本报告不执行或验证脚本。";
                }
            }
            catch (Exception ex) when (ex is FormatException or OverflowException)
            {
                code = "formula.unsupported-syntax"; feature = "formula/syntax";
                message = "当前公式语法或名称绑定不支持：" + ex.Message + " 结构编辑仍受现有安全检查限制。";
            }
            if (message is null) continue;
            message += cell.Calculation?.Value is null ? " 当前无缓存结果。" :
                cell.Calculation.State is CalculationState.Fresh or CalculationState.Error && cell.Calculation.InputRevision == snapshot.Revision
                    ? " 当前已有计算结果，但这不证明内置函数兼容性。" : " 当前显示值来自已有缓存，尚未确认是本次输入的计算结果。";
            diagnostics.Add(new(code, CompatibilityKind.CannotCalculate, feature, message,
                SheetName: sheet.Name, Range: TypedWorkbookXml.Reference(position)));
            formulaLocations++;
        }
        if (truncated) diagnostics.Add(new("report.formulas-truncated", CompatibilityKind.CannotCalculate, "报告容量",
            "已达到 1024 个公式问题位置；后续公式未逐项检查，未列出不代表可计算。"));
        if (richCount > 0) diagnostics.Add(new("text.rich-preserved", CompatibilityKind.PreservedOnly, "富文本",
            $"{richCount} 格保留分段字体和颜色；普通编辑需先显式转换为纯文本，转换可撤销。完整分段编辑尚未提供。"));
        if (richWarnings == 1024) diagnostics.Add(new("report.rich-display-truncated", CompatibilityKind.PreservedOnly, "报告容量", "富文本显示差异仅列出前 1024 格；未列出不代表完全对齐。"));
        var links = snapshot.Sheets.Sum(s => s.Annotations.Links.Count); var notes = snapshot.Sheets.Sum(s => s.Annotations.Notes.Count);
        if (links + notes > 0) diagnostics.Add(new("text.attachments-preserved", CompatibilityKind.PreservedOnly, "超链接和传统批注",
            $"{links} 个链接、{notes} 条传统批注可查看并保存；链接不自动访问。暂拒绝结构变换/排序，批注编辑和线程评论不在此支持范围内。"));
        if (formulaCount > 0) diagnostics.Add(new("formula.cache-summary", snapshot.FileState?.FormulaReadOnly == true ? CompatibilityKind.PreservedOnly : CompatibilityKind.Editable, "公式源码与缓存",
            $"共 {formulaCount} 个公式，{cachedCount} 个有缓存，{staleCount} 个缺少当前有效结果。能力检查不执行计算，也不证明缓存数值正确或引用这些公式的其他单元格可以正确重算。"));
        return new(unchanged, diagnostics);
    }
}
