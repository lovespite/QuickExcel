using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

public sealed class ScriptWorkerActionExecutor(string? executablePath = null) : IScriptActionExecutor
{
    private readonly string _executable = Path.GetFullPath(executablePath ?? Path.Combine(AppContext.BaseDirectory, "ScriptWorker", OperatingSystem.IsWindows() ? "QuickExcel.ScriptWorker.exe" : "QuickExcel.ScriptWorker"));
    private ScriptWorkerRuntime Create(ScriptActionRequest request, CancellationToken token) => ScriptWorkerRuntime.Start(_executable,
        new ScriptRuntimeRequest(request.Snapshot.ScriptModules, request.Stamp, TimeSpan.FromSeconds(5), token), actionsMode: true);

    public IReadOnlyList<ScriptActionDescriptor> List(ScriptActionRequest request, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, cancellationToken);
        using var runtime = Create(request, linked.Token);
        return runtime.ActionDescriptors;
    }

    public ScriptActionResult Run(ScriptActionRequest request, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, cancellationToken);
        linked.CancelAfter(TimeSpan.FromSeconds(5));
        linked.Token.ThrowIfCancellationRequested();
        if (request.Name is null) throw new InvalidOperationException("Missing action name.");
        if (request.Snapshot.Sheets.Sum(s => (long)s.Cells.Count) > ActionProtocol.MaxSnapshotCells)
            throw new InvalidOperationException("动作快照最多支持 100000 个存储单元格。");
        var snapshot = new ActionSnapshot(request.Stamp.InputRevision, request.Snapshot.DateSystem.ToString(), request.Snapshot.Sheets.Select(s => new ActionSheet(s.Id.ToString(), s.Name,
            s.Cells.OrderBy(c => c.Key.Row).ThenBy(c => c.Key.Column).Select(c => new ActionCell(c.Key.Row, c.Key.Column, ToWire(c.Value.Value),
                c.Value.Content.Kind == CellContentKind.Formula ? c.Value.Content.FormulaText : null,
                c.Value.Content.Kind == CellContentKind.Formula ? (c.Value.Calculation?.State.ToString() ?? "Dirty") : "literal",
                c.Value.Content.Kind == CellContentKind.Literal || c.Value.Calculation?.Value is not null)).ToArray())).ToArray());
        ActionProtocol.ValidateSnapshot(snapshot);
        using var runtime = Create(request, linked.Token);
        if (!runtime.ActionDescriptors.Any(a => a.Name == request.Name))
        {
            if (request.Event is not null) return new(request.OperationId, request.Stamp.InputRevision, request.Name, [], []);
            throw new InvalidOperationException("找不到已绑定的操作脚本。");
        }
        var patch = runtime.RunAction(new(request.OperationId, request.Name, snapshot)
        {
            Files = request.Files.Entries.Values.Select(e => new VirtualEntry(e.Path, e.IsDirectory ? null : Convert.ToBase64String(e.Data.AsSpan()), e.IsDirectory)).ToArray(),
            Event = request.Event is { } ev ? new(ev.Name, ev.SessionId, ev.Sequence, ev.Revision, ev.SheetId, ev.SheetName, ev.Ranges.ToArray(), ev.IsUndoOrRedo) : null
        });
        var logs = runtime.GetLogs().ToImmutableArray();
        linked.Token.ThrowIfCancellationRequested();
        return new(patch.OperationId, patch.BaseRevision, patch.Name, patch.Edits.Select(e => new ScriptActionEdit(
            new(new SheetId(Guid.Parse(e.SheetId)), e.Row, e.Column), e.Value is null ? null : FromWire(e.Value), e.Formula)).ToImmutableArray(), logs)
        { FileEdits = patch.FileEdits.Select(e => new VirtualFileEdit(e.Operation, e.Path, e.Base64 is null ? [] : VirtualFileProtocol.Bytes(e.Base64).ToImmutableArray())).ToImmutableArray() };
    }

    private static WireValue ToWire(CellValue value) => value.Kind switch
    {
        CellValueKind.Error => new("error", JsonSerializer.SerializeToElement(value.ErrorCode.ToString())),
        CellValueKind.Blank => WireValue.FromObject(null),
        CellValueKind.Number => WireValue.FromObject(value.NumberValue),
        CellValueKind.Text => WireValue.FromObject(value.TextValue),
        CellValueKind.Boolean => WireValue.FromObject(value.BooleanValue),
        _ => throw new InvalidOperationException("Unsupported snapshot value.")
    };
    private static CellValue FromWire(WireValue value)
    {
        ActionProtocol.Scalar(value);
        return value.Kind switch
        {
            "blank" => CellValue.Blank,
            "number" => CellValue.FromNumber(value.Value.GetDouble()),
            "text" => CellValue.FromText(value.Value.GetString()!),
            "boolean" => CellValue.FromBoolean(value.Value.GetBoolean()),
            "error" => CellValue.FromError(Enum.Parse<CellErrorCode>(value.Value.GetString()!)),
            _ => throw new ProtocolException("Unsupported action value.")
        };
    }
}
