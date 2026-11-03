using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed record ScriptActionDescriptor(string Name, string Description);
public sealed record ScriptActionRequest(CalculationStamp Stamp, Guid OperationId, string? Name, WorkbookSnapshot Snapshot, CancellationToken CancellationToken)
{
    public VirtualFileSet Files { get; init; } = VirtualFileSet.Empty;
    public ScriptEvent? Event { get; init; }
    /// <summary>Only an explicit session console request sets this; workbook modules are never loaded for it.</summary>
    public string? ConsoleSource { get; internal init; }
}
public sealed record ScriptActionEdit(CellAddress Address, CellValue? Value, string? Formula);
public sealed record ScriptActionResult(Guid OperationId, long BaseRevision, string Name, ImmutableArray<ScriptActionEdit> Edits, ImmutableArray<string> Logs)
{
    public ImmutableArray<VirtualFileEdit> FileEdits { get; init; } = [];
}
public interface IScriptActionExecutor
{
    IReadOnlyList<ScriptActionDescriptor> List(ScriptActionRequest request, CancellationToken cancellationToken);
    ScriptActionResult Run(ScriptActionRequest request, CancellationToken cancellationToken);
}

/// <summary>Only the session's authorized action commit can create this command.</summary>
internal sealed class ApplyActionPatchCommand(ScriptActionRequest request, ScriptActionResult result, IWorkbookVirtualFiles? files = null) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (result.OperationId != request.OperationId || result.BaseRevision != request.Stamp.InputRevision || result.Name != request.Name ||
            result.Edits.IsDefault || result.Edits.Length > 20_000 || result.Logs.IsDefault || result.Logs.Length > 100 || result.Logs.Any(l => l is null || l.Length > 4000))
            throw new InvalidOperationException("动作结果身份或容量无效，变更未提交。");
        if (result.FileEdits.IsDefault) throw new InvalidOperationException("缺少虚拟文件变更。");
        if (request.Event?.Name == "workbook.close" && (result.Edits.Length != 0 || result.FileEdits.Length != 0))
            throw new InvalidOperationException("关闭完成事件只读；保存前的清理请使用 workbook.beforeClose。");
        if (request.Event?.Name is "workbook.calculated" or "workbook.saved" && (result.Edits.Length != 0 || result.FileEdits.Any(e => e.Path.StartsWith("//xlsq/", StringComparison.Ordinal))))
            throw new InvalidOperationException("完成通知不能修改持久输入；请使用 beforeSave 或其他交互事件。");
        var packageEdits = result.FileEdits.Where(e => e.Path.StartsWith("//xlsq/", StringComparison.Ordinal)).ToArray();
        if (!result.Edits.IsEmpty && packageEdits.Any(e => e.Path == "//xlsq/workbook.xlsx"))
            throw new InvalidOperationException("不能在同一事务混用 workbook.xlsx 原始替换与单元格 patch。");
        var package = packageEdits.Length == 0 ? snapshot : (files ?? throw new NotSupportedException("未配置包文件适配器。")).Apply(snapshot, packageEdits);
        if (!ReferenceEquals(package.ScriptModules, snapshot.ScriptModules))
        {
            if (request.ConsoleSource is not null)
                throw new InvalidOperationException("控制台不能改写脚本源码、绑定或签名；请使用脚本编辑器。控制台的数据修改须保持可撤销。");
            var sourcePaths = snapshot.ScriptModules.Modules.Values.Concat(package.ScriptModules.Modules.Values)
                .Select(m => "//xlsq/" + m.Path).Concat(["//xlsq/manifest.json", "//xlsq/state/document.json", "//xlsq/security/script-signature.json"])
                .ToHashSet(StringComparer.Ordinal);
            if (!result.Edits.IsEmpty || packageEdits.Any(e => !sourcePaths.Contains(e.Path)))
                throw new InvalidOperationException("源码、绑定或签名修改必须独立于单元格和其他永久文件修改提交。");
        }
        var edits = new List<CellEdit>();
        foreach (var edit in result.Edits)
        {
            if (edit is null || (edit.Value is null) == (edit.Formula is null)) throw new InvalidOperationException("动作必须指定值或公式。");
            WorkbookLimits.ValidatePosition(edit.Address.Row, edit.Address.Column);
            var sheet = snapshot.GetSheet(edit.Address.SheetId);
            // Actions change content, never styles, merged cells or image anchors.
            if (!sheet.MergedCells.IsEmpty || sheet.Images.Any(i => i.PlaceInCell))
                throw new InvalidOperationException("动作暂不支持写入含合并区域或单元格内图片的工作表。");
            var style = sheet.GetStyleId(edit.Address.Row, edit.Address.Column);
            if (edit.Formula is { } formula && (formula.Length is < 2 or > 8192 || !formula.StartsWith('=')))
                throw new InvalidOperationException("动作公式必须以 = 开头且不超过 8192 字符。");
            if (edit.Formula is { } syntax) _ = FormulaSyntax.Parse(syntax);
            edits.Add(new(edit.Address, edit.Formula is { } f ? CellContent.FromFormula(f, style) : CellContent.FromLiteral(edit.Value!.Value, style)));
        }
        var draft = new SetCellsCommand(edits).Prepare(package);
        if (ReferenceEquals(package, snapshot)) return draft;
        var changes = new List<CellChange>();
        foreach (var sheet in snapshot.Sheets.Concat(draft.Snapshot.Sheets).DistinctBy(s => s.Id))
        {
            var before = snapshot.Sheets.FirstOrDefault(s => s.Id == sheet.Id);
            var after = draft.Snapshot.Sheets.FirstOrDefault(s => s.Id == sheet.Id);
            foreach (var position in (before?.Cells.Keys ?? []).Union(after?.Cells.Keys ?? []))
            {
                var a = before?.GetCell(position.Row, position.Column); var b = after?.GetCell(position.Row, position.Column);
                if (a?.Content != b?.Content) changes.Add(new(new(sheet.Id, position.Row, position.Column), a, b));
            }
        }
        return new(draft.Snapshot, changes.ToImmutableArray(), true);
    }
}
