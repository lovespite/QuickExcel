using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed record ScriptActionDescriptor(string Name, string Description);
public sealed record ScriptActionRequest(CalculationStamp Stamp, Guid OperationId, string? Name, WorkbookSnapshot Snapshot, CancellationToken CancellationToken);
public sealed record ScriptActionEdit(CellAddress Address, CellValue? Value, string? Formula);
public sealed record ScriptActionResult(Guid OperationId, long BaseRevision, string Name, ImmutableArray<ScriptActionEdit> Edits, ImmutableArray<string> Logs);
public interface IScriptActionExecutor
{
    IReadOnlyList<ScriptActionDescriptor> List(ScriptActionRequest request, CancellationToken cancellationToken);
    ScriptActionResult Run(ScriptActionRequest request, CancellationToken cancellationToken);
}

/// <summary>Only the session's authorized action commit can create this command.</summary>
internal sealed class ApplyActionPatchCommand(ScriptActionRequest request, ScriptActionResult result) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (result.OperationId != request.OperationId || result.BaseRevision != request.Stamp.InputRevision || result.Name != request.Name ||
            result.Edits.IsDefault || result.Edits.Length > 20_000 || result.Logs.IsDefault || result.Logs.Length > 100 || result.Logs.Any(l => l is null || l.Length > 4000))
            throw new InvalidOperationException("动作结果身份或容量无效，变更未提交。");
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
        return new SetCellsCommand(edits).Prepare(snapshot);
    }
}
