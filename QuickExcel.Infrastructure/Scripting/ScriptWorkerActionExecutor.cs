using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Runtime;

public sealed class ScriptWorkerActionExecutor(string? executablePath = null) : IScriptActionExecutor
{
    private readonly string _executable = Path.GetFullPath(executablePath ?? Path.Combine(AppContext.BaseDirectory, "ScriptWorker", OperatingSystem.IsWindows() ? "QuickExcel.ScriptWorker.exe" : "QuickExcel.ScriptWorker"));
    private static ScriptModuleSet ExecutionModules(ScriptActionRequest request)
    {
        if (request.ConsoleSource is not { } source) return request.Snapshot.ScriptModules;
        // Serialize the code as data. Direct eval runs only inside the normal
        // zero-capability action worker, with this immutable transaction context.
        // The console is one synthetic main(workbook, host) entry. Direct eval runs
        // only inside the zero-capability action worker with this transaction context.
        var text = "export function main(workbook, host) { const result = eval(" +
            JsonSerializer.Serialize(source, WorkerProtocol.TypeInfo<string>()) + "); console.log('=>', result); }";
        return new(new Dictionary<string, ScriptModule> { ["__console__"] = new("scripts/console.js", text) }, ["__console__"],
            scriptApiVersion: 4, bindings: [new("__console__", "actions")]);
    }
    private ScriptWorkerLease Create(ScriptActionRequest request, CancellationToken token, TimeSpan? timeLimit = null) => SharedScriptWorkers.Acquire(_executable,
        new ScriptRuntimeRequest(ExecutionModules(request), request.Stamp, timeLimit ?? TimeSpan.FromSeconds(5), token) { RuntimeLifetime = request.RuntimeLifetime }, actions: true);

    public IReadOnlyList<ScriptActionDescriptor> List(ScriptActionRequest request, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, cancellationToken);
        IReadOnlyList<ScriptActionDescriptor> actions;
        IReadOnlyList<string> hosts;
        using (var lease = Create(request, linked.Token))
        {
            actions = lease.Runtime.ActionDescriptors;
            hosts = lease.Runtime.NetworkHosts;
        }
        // The user may take arbitrarily long to decide. End the worker round
        // before prompting, keeping the isolated context without an active deadline.
        request.Network.EnsureAuthorized(hosts, linked.Token);
        return actions;
    }

    public ScriptActionResult Run(ScriptActionRequest request, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken, cancellationToken);
        var networkOptions = request.Network.Options;
        var isAction = request.ConsoleSource is null && request.Event is null && request.Name is { } name && !name.StartsWith('@') &&
            request.Snapshot.ScriptModules.EntriesFor("actions").Contains(name);
        var timeLimit = isAction ? networkOptions.ActionTimeout : TimeSpan.FromSeconds(5);
        ScriptCompilation.PrepareAsync(ExecutionModules(request), request.Stamp, linked.Token).GetAwaiter().GetResult();
        IReadOnlyList<string> authorizedHosts = [];
        if (request.ConsoleSource is null)
        {
            using (var discovery = Create(request, linked.Token)) authorizedHosts = discovery.Runtime.NetworkHosts.ToArray();
            request.Network.EnsureAuthorized(authorizedHosts, linked.Token);
        }
        linked.CancelAfter(timeLimit);
        linked.Token.ThrowIfCancellationRequested();
        if (request.Name is null) throw new InvalidOperationException("Missing action name.");
        if (OperatingSystem.IsMacOS() && request.Model is not null && request.RuntimeLifetime.CanBeCanceled &&
            request.ConsoleSource is null && request.Snapshot.ScriptModules.Bindings.Any(b => b.Kind == "formula-functions"))
        {
            // A live model read can calculate UDFs, including during v4 initialization.
            // macOS's durable private-payload copy can exceed the native action's
            // one-second budget. Prepare the separate persistent formula context
            // before entering that operation, within the same five-second host
            // deadline. No UDF is invoked and no calculation result is published.
            using var formulas = SharedScriptWorkers.Acquire(_executable,
                new ScriptRuntimeRequest(request.Snapshot.ScriptModules, request.Stamp, TimeSpan.FromSeconds(5), linked.Token)
                { RuntimeLifetime = request.RuntimeLifetime }, actions: false);
        }
        if (request.Model is null && request.Snapshot.Sheets.Sum(s => (long)s.Cells.Count) > ActionProtocol.MaxSnapshotCells)
            throw new InvalidOperationException("动作快照最多支持 100000 个存储单元格。");
        var snapshot = new ActionSnapshot(request.Stamp.InputRevision, request.Snapshot.DateSystem.ToString(), request.Snapshot.Sheets.Select(s => new ActionSheet(s.Id.ToString(), s.Name,
            request.Model is not null ? [] : s.Cells.OrderBy(c => c.Key.Row).ThenBy(c => c.Key.Column).Select(c => new ActionCell(c.Key.Row, c.Key.Column, ToWire(c.Value.Value),
                c.Value.Content.Kind == CellContentKind.Formula ? c.Value.Content.FormulaText : null,
                c.Value.Content.Kind == CellContentKind.Formula ? (c.Value.Calculation?.State.ToString() ?? "Dirty") : "literal",
                c.Value.Content.Kind == CellContentKind.Literal || c.Value.Calculation?.Value is not null)).ToArray())).ToArray());
        ActionProtocol.ValidateSnapshot(snapshot);
        using var lease = Create(request, linked.Token, timeLimit);
        var runtime = lease.Runtime;
        if (request.ConsoleSource is null && !runtime.NetworkHosts.SequenceEqual(authorizedHosts))
            throw new InvalidOperationException("网络权限声明在加载期间发生变化，请重新应用脚本。");
        if (!runtime.ActionDescriptors.Any(a => a.Name == request.Name))
        {
            if (request.Event is not null) return new(request.OperationId, request.Stamp.InputRevision, request.Name, [], []);
            throw new InvalidOperationException("找不到已绑定的操作脚本。");
        }
        var adapter = request.Model is null ? null : new ScriptModelAdapter(request, linked.Token);
        var networkRequests = 0;
        ModelReply HostCall(ModelCall call)
        {
            if (call.Request.ValueKind != JsonValueKind.Object ||
                !call.Request.TryGetProperty("method", out var method) || method.GetString() != "network.fetch")
                return adapter?.Call(call) ?? throw new ProtocolException("No live model for this request.");
            linked.Token.ThrowIfCancellationRequested();
            request.RuntimeLifetime.ThrowIfCancellationRequested();
            if (call.OperationId != request.OperationId || call.Request.EnumerateObject().Count() != 2 ||
                !call.Request.TryGetProperty("args", out var args) || args.ValueKind != JsonValueKind.Array || args.GetArrayLength() != 1)
                throw new ProtocolException("Invalid network request.");
            try
            {
                // Check the host-owned action identity, independently of the JS gate.
                if (request.ConsoleSource is not null || request.Event is not null || request.Name.StartsWith('@') ||
                    !request.Snapshot.ScriptModules.EntriesFor("actions").Contains(request.Name))
                    throw new InvalidOperationException("fetch 只能在 Actions 中调用。");
                if (++networkRequests > 16) throw new InvalidOperationException("一次 Action 最多发送 16 个网络请求。");
                using var access = CancellationTokenSource.CreateLinkedTokenSource(linked.Token, request.Network.CaptureNetworkAccess(authorizedHosts));
                var response = ScriptNetworkBroker.Fetch(args[0], authorizedHosts, access.Token, request.Network.Record, networkOptions);
                access.Token.ThrowIfCancellationRequested();
                return new(call.OperationId, call.Sequence, true, response, null);
            }
            catch (Exception error) when (error is ArgumentException or InvalidOperationException or IOException or HttpRequestException)
            {
                var message = error.Message.Length <= 1000 ? error.Message : "网络请求失败。";
                return new(call.OperationId, call.Sequence, false, WorkerProtocol.EncodePayload<string?>(null), message) { Code = "NetworkError" };
            }
        }
        var patch = runtime.RunAction(new(request.OperationId, request.Name, snapshot)
        {
            LiveModel = request.Model is not null,
            MaxNetworkBodyBytes = networkOptions.BodyLimitBytes,
            AllowNetwork = request.ConsoleSource is null && request.Event is null && !request.Name.StartsWith('@') &&
                request.Snapshot.ScriptModules.EntriesFor("actions").Contains(request.Name),
            Files = request.Model is not null ? [] : request.Files.Entries.Values.Select(e => new VirtualEntry(e.Path, e.IsDirectory ? null : Convert.ToBase64String(e.Data.AsSpan()), e.IsDirectory)).ToArray(),
            Event = request.Event is { } ev ? new(ev.Name, ev.SessionId, ev.Sequence, ev.Revision, ev.SheetId, ev.SheetName, ev.Ranges.ToArray(), ev.IsUndoOrRedo) : null
        }, HostCall);
        var logs = runtime.GetLogs().ToImmutableArray();
        linked.Token.ThrowIfCancellationRequested();
        return new(patch.OperationId, patch.BaseRevision, patch.Name, patch.Edits.Select(e => new ScriptActionEdit(
            new(new SheetId(Guid.Parse(e.SheetId)), e.Row, e.Column), e.Value is null ? null : FromWire(e.Value), e.Formula)).ToImmutableArray(), logs)
        { FileEdits = patch.FileEdits.Select(e => new VirtualFileEdit(e.Operation, e.Path, e.Base64 is null ? [] : VirtualFileProtocol.Bytes(e.Base64).ToImmutableArray())).ToImmutableArray() };
    }

    private static WireValue ToWire(CellValue value) => value.Kind switch
    {
        CellValueKind.Error => new("error", JsonSerializer.SerializeToElement(value.ErrorCode.ToString(), WorkerProtocol.TypeInfo<string>())),
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
