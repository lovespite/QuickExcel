using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

public sealed record FormulaAuditTarget(CellAddress Address, CellRectangle Range, string Label, string State, string Value, string Formula);
public sealed record FormulaAuditResult(CellAddress Origin, string Formula, string State, string Value,
    IReadOnlyList<FormulaAuditTarget> Precedents, IReadOnlyList<FormulaAuditTarget> Dependents, IReadOnlyList<string> Notices, bool Complete);

/// <summary>A bounded, read-only view of the calculator's static value-dependency index. Never executes scripts.</summary>
public static class FormulaAudit
{
    public static FormulaAuditResult Inspect(WorkbookSnapshot snapshot, CellAddress origin, CancellationToken cancellationToken = default, int limit = 500)
    {
        if (limit is < 1 or > 5000) throw new ArgumentOutOfRangeException(nameof(limit));
        WorkbookLimits.ValidatePosition(origin.Row, origin.Column); var sheet = snapshot.GetSheet(origin.SheetId);
        var work = 0;
        void Check() { cancellationToken.ThrowIfCancellationRequested(); if (++work > 10_000_000) throw new InvalidOperationException(MessageText.Get("公式审核超过 10,000,000 步预算，请缩小工作簿后重试。")); }
        var spills = new Dictionary<CellAddress, FormulaSpill>();
        foreach (var s in snapshot.Sheets)
        foreach (var (position, cell) in s.Cells)
        {
            Check();
            if (cell.Content.Kind == CellContentKind.Formula && cell.Calculation is { Array: { } array, State: CalculationState.Fresh or CalculationState.Error } calculation && calculation.InputRevision == snapshot.Revision)
            { var address = new CellAddress(s.Id, position.Row, position.Column); spills[address] = new(address, array); }
        }
        var index = new WorkbookDependencyIndex(snapshot, Check, spills); var notices = new List<string>(); var complete = true;
        void Notice(string text) { if (!notices.Contains(text)) notices.Add(text); complete = false; }
        bool Incomplete(WorkbookDependencyIndex.Node node) => node.ParseError is not null || node.Script || node.VisibilityBlocked ||
            FormulaSyntax.Walk(node.Expression).Any(p => p is ValueNode { Value.Kind: CellValueKind.Error, Value.ErrorCode: CellErrorCode.Reference or CellErrorCode.Name });
        if (index.Nodes.Values.Any(Incomplete)) Notice(MessageText.Get("工作簿含未完整分析的公式或脚本；列表仅表示已知静态值依赖，可能不完整。"));
        if (snapshot.FileState?.FormulaReadOnly == true) Notice(MessageText.Get("此来源工作簿含未支持的公式语义，依赖列表仅供参考。"));
        var precedents = new List<FormulaAuditTarget>(); var dependents = new List<FormulaAuditTarget>();
        var originCell = sheet.GetCell(origin.Row, origin.Column);
        if (originCell?.SpillOwner is { } owner)
        {
            precedents.Add(Target(snapshot, owner, new(owner.Row, owner.Column, owner.Row, owner.Column)));
            notices.Add(MessageText.Get("当前单元格是溢出结果；直接来源列表显示数组锚点。"));
        }
        else if (index.Nodes.TryGetValue(origin, out var node))
        {
            if (ProtectionSafety.HidesFormula(snapshot, sheet, origin.Row, origin.Column)) Notice(MessageText.Get("受保护公式的源码和依赖关系不显示。"));
            else
            {
                if (node.ParseError is not null) Notice(MessageText.Get("公式无法完整绑定：") + node.ParseError);
                var seen = new HashSet<(SheetId, CellRectangle)>();
                foreach (var reference in FormulaSyntax.Walk(ConditionalRanges.ForDependencies(node.Expression)).Concat(SubtotalFunctions.FilterDependencies(node.Expression, snapshot)).OfType<ReferenceNode>())
                {
                    Check(); var targetSheet = snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
                    if (targetSheet is null) { Notice(MessageText.Get("存在缺失的引用目标。")); continue; }
                    var range = new CellRectangle(reference.Row, reference.Column, Math.Min(WorkbookLimits.MaxRows - 1, reference.EndRow), Math.Min(WorkbookLimits.MaxColumns - 1, reference.EndColumn));
                    if (reference.EndRow >= WorkbookLimits.MaxRows || reference.EndColumn >= WorkbookLimits.MaxColumns)
                        notices.Add(MessageText.Get("整列或超出网格的引用按当前容量范围定位。"));
                    if (!seen.Add((targetSheet.Id, range))) continue;
                    if (precedents.Count == limit) { Notice(MessageText.Get("审核列表达到显示上限，请缩小审核范围。")); break; }
                    precedents.Add(Target(snapshot, new(targetSheet.Id, range.Top, range.Left), range));
                }
            }
        }
        IEnumerable<CellAddress> Consumers()
        {
            foreach (var address in index.DirectDependents(origin, Check)) yield return address;
            // A formula reading a spill child depends on its anchor even without a literal anchor reference.
            if (index.Nodes.TryGetValue(origin, out var node)) foreach (var address in node.Dependents) yield return address;
        }
        foreach (var address in Consumers().Distinct().OrderBy(a => snapshot.Sheets.IndexOf(snapshot.GetSheet(a.SheetId))).ThenBy(a => a.Row).ThenBy(a => a.Column))
        {
            Check(); var targetSheet = snapshot.GetSheet(address.SheetId);
            if (ProtectionSafety.HidesFormula(snapshot, targetSheet, address.Row, address.Column)) { Notice(MessageText.Get("受保护公式的源码和依赖关系不显示。")); continue; }
            if (dependents.Count == limit) { Notice(MessageText.Get("审核列表达到显示上限，请缩小审核范围。")); break; }
            dependents.Add(Target(snapshot, address, new(address.Row, address.Column, address.Row, address.Column)));
        }
        if (index.CycleMembers.Contains(origin)) notices.Add(MessageText.Get("当前公式处于循环依赖中。"));
        var current = Target(snapshot, origin, new(origin.Row, origin.Column, origin.Row, origin.Column));
        return new(origin, current.Formula, current.State, current.Value, precedents, dependents, notices.Distinct().ToArray(), complete);
    }

    public static FormulaAuditTarget Target(WorkbookSnapshot snapshot, CellAddress address, CellRectangle range)
    {
        var sheet = snapshot.GetSheet(address.SheetId); var cell = sheet.GetCell(address.Row, address.Column);
        var formula = cell?.Content.Kind == CellContentKind.Formula;
        var calculated = formula || cell?.SpillOwner is not null;
        var fresh = cell?.Calculation is { Value: not null, State: CalculationState.Fresh or CalculationState.Error } calculation && calculation.InputRevision == snapshot.Revision;
        var state = !calculated ? MessageText.Get("常量 / 空白") : cell?.Calculation is null ? MessageText.Get("未计算") : !fresh ? MessageText.Get("旧缓存 / 待重算") :
            cell.Calculation.State == CalculationState.Error ? MessageText.Get("当前错误") : MessageText.Get("当前结果");
        if (calculated && !fresh && cell?.Calculation is { } old) state += " · " + MessageText.Get(old.State switch { CalculationState.Running => "正在计算", CalculationState.Blocked => "已阻止", CalculationState.Dirty => "已改变", _ => "过期" });
        var value = range.Count == 1 ? CellInput.DisplayText(cell?.Value ?? CellValue.Blank) : MessageText.Format($"区域：{range.Rows:N0} 行 × {range.Columns:N0} 列");
        if (value.Length > 512) value = value[..512] + "…";
        var label = "'" + sheet.Name.Replace("'", "''") + "'!" + (range.Count == 1 ? range.ToString().Split(':')[0] : range.ToString());
        return new(address, range, label, range.Count == 1 ? state : MessageText.Get("范围引用"), value,
            formula && !ProtectionSafety.HidesFormula(snapshot, sheet, address.Row, address.Column) ? cell!.Content.FormulaText : "");
    }
}
