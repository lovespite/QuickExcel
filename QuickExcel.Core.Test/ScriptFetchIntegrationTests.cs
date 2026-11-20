using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptFetchIntegrationTests
{
    private static WorkbookDocument Document(string action, string initialization = "host.permissions.network.request('127.0.0.1');", string? events = null)
    {
        var modules = new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", initialization),
            ["download"] = new("scripts/download.js", action)
        };
        var bindings = new List<ScriptBinding> { new("main", "initialization"), new("download", "actions") };
        if (events is not null) { modules.Add("events", new("scripts/events.js", events)); bindings.Add(new("events", "events")); }
        return new(Book(Sheet("Data")).WithScriptModules(new(modules, modules.Keys, scriptApiVersion: 4, bindings: bindings)),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
    }

    [Fact]
    public async Task AuthorizationPrecedesInitializationAndAsyncFetchCommitsOnce()
    {
        await using var server = new ScriptFetchTestServer(async _ =>
        {
            await Task.Delay(1200); // Network waiting must not consume the one-second JS watchdog.
            return new(200, "{\"answer\":42}", new Dictionary<string, string> { ["Content-Type"] = "application/json" });
        });
        using var doc = Document($$"""
            export async function main(w) {
              const r = await fetch('{{server.Url}}answer');
              const copy = r.clone();
              if (!r.ok || r.status !== 200 || !r.headers.get('CONTENT-TYPE').includes('json')) throw Error('response');
              const data = await r.json();
              if (!r.bodyUsed || (await copy.text()) !== '{"answer":42}') throw Error('body');
              let used = false; try { await r.text(); } catch(e) { used = e instanceof TypeError; }
              if (!used) throw Error('must consume once');
              w.getActiveWorksheet().getRange('A1').setValue(data.answer);
              await Promise.resolve();
              w.getActiveWorksheet().getRange('B1').setValue('done');
            }
            """, "host.permissions.network.request('127.0.0.1'); export async function main(w) { await Promise.resolve(); w.getActiveWorksheet().getRange('C1').setValue('initialized'); }");
        var prompts = 0;
        doc.Session.Network.Authorize = (hosts, _) =>
        {
            prompts++; Assert.Equal(new[] { "127.0.0.1" }, hosts); Assert.Equal("", doc.GetCellValue(0, 0, 2)); return true;
        };
        doc.Session.Network.RecordingEnabled = true;
        doc.EnableScriptsAndCalculate(); doc.ListActions();
        Assert.Equal("initialized", doc.GetCellValue(0, 0, 2));
        doc.RunAction("download");
        Assert.Equal(1, prompts); Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Equal("done", doc.GetCellValue(0, 0, 1));
        var row = Assert.Single(doc.Session.Network.Entries); Assert.Equal(200, row.Status); Assert.NotEmpty(row.ResponseBody!);
        Assert.True(doc.Undo()); Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.Equal("", doc.GetCellValue(0, 0, 1));
        Assert.Equal("initialized", doc.GetCellValue(0, 0, 2));
    }

    [Fact]
    public void RepeatedLoadingAndSourceEditsReuseUnchangedNetworkConsentUntilExplicitDisable()
    {
        const string initialize = "export function main(w) { const cell = w.getActiveWorksheet().getRange('C1'); cell.setValue(Number(cell.getValue() || 0) + 1); }";
        using var doc = Document("export function main(w) { w.getActiveCell().setValue('ready'); }",
            "host.permissions.network.request('api.example.com'); host.permissions.network.request('127.0.0.1'); " + initialize);
        var prompts = new List<string[]>();
        doc.Session.Network.Authorize = (hosts, _) => { prompts.Add(hosts.ToArray()); return true; };

        doc.EnableScriptsAndCalculate();
        doc.EnableScriptsAndCalculate();
        doc.ListActions();
        Assert.Single(prompts); Assert.Equal(new[] { "127.0.0.1", "api.example.com" }, prompts[0]);
        Assert.Equal("2", doc.GetCellValue(0, 0, 2));

        // Editing and loading still replace the execution context, while the complete
        // normalized network declaration remains the same despite order and duplicates.
        doc.ScriptSource = "host.permissions.network.request('127.0.0.1'); host.permissions.network.request('API.EXAMPLE.COM.'); host.permissions.network.request('api.example.com'); " + initialize;
        Assert.False(doc.ScriptsEnabled);
        doc.EnableScriptsAndCalculate(); doc.RunAction("download");
        Assert.Single(prompts); Assert.Equal("3", doc.GetCellValue(0, 0, 2));
        Assert.Equal("ready", doc.GetCellValue(0, 0, 0));

        doc.DisableScripts();
        doc.EnableScriptsAndCalculate();
        Assert.Equal(2, prompts.Count); Assert.Equal("4", doc.GetCellValue(0, 0, 2));
    }

    [Fact]
    public void FailedInitializationKeepsConsentButNeverReplaysOrRetainsTheFailedContext()
    {
        const string declaration = "host.permissions.network.request('api.example.com'); ";
        using var doc = Document("export function main() {}", declaration +
            "export function main(w) { w.getActiveCell().setValue('partial'); throw Error('initialization failed'); }");
        var prompts = 0;
        doc.Session.Network.Authorize = (_, _) => { prompts++; return true; };
        Assert.ThrowsAny<Exception>(() => doc.EnableScriptsAndCalculate());
        Assert.False(doc.ScriptsEnabled); Assert.Equal("", doc.GetCellValue(0, 0, 0));
        Assert.ThrowsAny<Exception>(() => doc.EnableScriptsAndCalculate());
        Assert.False(doc.ScriptsEnabled); Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.Equal(1, prompts);
        doc.ScriptSource = declaration + "export function main(w) { w.getActiveCell().setValue('initialized'); }";
        doc.EnableScriptsAndCalculate();
        Assert.True(doc.ScriptsEnabled); Assert.Equal("initialized", doc.GetCellValue(0, 0, 0)); Assert.Equal(1, prompts);
    }

    [Fact]
    public void ChangedNetworkListRequiresFreshConsentBeforeInitialization()
    {
        using var doc = Document("export function main() {}",
            "host.permissions.network.request('api.example.com'); host.permissions.network.request('127.0.0.1'); export function main(w) { w.getActiveCell().setValue('original'); }");
        var prompts = new List<string[]>(); var allow = true;
        doc.Session.Network.Authorize = (hosts, _) =>
        {
            prompts.Add(hosts.ToArray());
            if (prompts.Count > 1) Assert.Equal("original", doc.GetCellValue(0, 0, 0));
            return allow;
        };
        doc.EnableScriptsAndCalculate();

        // Removing a host changes the complete list too; an earlier superset grant
        // cannot silently authorize this new declaration or run its initialization.
        doc.ScriptSource = "host.permissions.network.request('api.example.com'); export function main(w) { w.getActiveCell().setValue('changed'); }";
        allow = false;
        Assert.ThrowsAny<Exception>(() => doc.EnableScriptsAndCalculate());
        Assert.False(doc.ScriptsEnabled); Assert.Equal("original", doc.GetCellValue(0, 0, 0));
        Assert.Equal(2, prompts.Count); Assert.Equal(new[] { "api.example.com" }, prompts[1]);
        Assert.Throws<InvalidOperationException>(() => doc.Session.Network.CaptureNetworkAccess(["127.0.0.1"]));

        allow = true;
        doc.EnableScriptsAndCalculate();
        Assert.Equal(3, prompts.Count); Assert.Equal("changed", doc.GetCellValue(0, 0, 0));
        Assert.Throws<InvalidOperationException>(() => doc.Session.Network.CaptureNetworkAccess(["127.0.0.1"]));
        doc.EnableScriptsAndCalculate(); Assert.Equal(3, prompts.Count);
    }

    [Fact]
    public async Task ReturnedPromiseAndHttpErrorStatusFollowFetchSemantics()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(404, "missing")));
        using var doc = Document($$"""
            export function main(w) {
              return fetch('{{server.Url}}missing').then(async r => {
                if (r.ok || r.status !== 404) throw Error('status');
                w.getActiveCell().setValue(await r.text());
              });
            }
            """);
        doc.Session.Network.Authorize = (_, _) => true;
        doc.EnableScriptsAndCalculate(); doc.RunAction("download");
        Assert.Equal("missing", doc.GetCellValue(0, 0, 0)); Assert.Empty(doc.Session.Network.Entries);
    }

    [Fact]
    public async Task UndeclaredHostAndRejectedGrantNeverReachServer()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        using (var denied = Document($"export async function main() {{ await fetch('{server.Url}'); }}",
            "host.permissions.network.request('127.0.0.1'); export function main(w) { w.getActiveCell().setValue(7); }"))
        {
            denied.Session.Network.Authorize = (_, _) => false;
            Assert.ThrowsAny<Exception>(() => denied.EnableScriptsAndCalculate());
            Assert.Equal("", denied.GetCellValue(0, 0, 0)); Assert.False(denied.ScriptsEnabled);
        }
        using (var undeclared = Document($"export async function main(w) {{ try {{ await fetch('{server.Url}'); }} catch(e) {{ w.getActiveCell().setValue(e instanceof TypeError); }} }}", ""))
        {
            undeclared.EnableScriptsAndCalculate(); undeclared.RunAction("download"); Assert.Equal("TRUE", undeclared.GetCellValue(0, 0, 0));
        }
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task RejectedAsyncActionRollsBackWritesAfterNetworkResponse()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        using var doc = Document($"export async function main(w) {{ w.getActiveCell().setValue(1); await fetch('{server.Url}'); w.getActiveWorksheet().getRange('B1').setValue(2); throw Error('after network'); }}");
        doc.Session.Network.Authorize = (_, _) => true; doc.EnableScriptsAndCalculate();
        Assert.ThrowsAny<Exception>(() => doc.RunAction("download"));
        Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.Equal("", doc.GetCellValue(0, 0, 1)); Assert.Single(server.Requests);
    }

    [Fact]
    public async Task InitializationAndEventsCannotFetchEvenAfterGrant()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        using var init = Document("export function main() {}",
            $"host.permissions.network.request('127.0.0.1'); export async function main() {{ await fetch('{server.Url}'); }}");
        init.Session.Network.Authorize = (_, _) => true;
        Assert.ThrowsAny<Exception>(() => init.EnableScriptsAndCalculate());
        using var events = Document("export function main() {}", events:
            $"export function onCellClick(c) {{ fetch('{server.Url}'); }}");
        events.Session.Network.Authorize = (_, _) => true; events.EnableScriptsAndCalculate();
        events.Session.RaiseScriptEvent("cell.click", events.Snapshot.ActiveSheetId);
        Assert.ThrowsAny<Exception>(() => events.DispatchScriptEvents());
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task ActionDeclarationCannotExpandInitializationPermissions()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        using var doc = Document("host.permissions.network.request('localhost'); export function main() {}", "");
        var prompts = 0; doc.Session.Network.Authorize = (_, _) => { prompts++; return true; };
        Assert.ThrowsAny<Exception>(() => doc.EnableScriptsAndCalculate());
        Assert.Equal(0, prompts); Assert.Empty(server.Requests);
    }

    [Fact]
    public void ConsentDialogWaitingDoesNotConsumeWorkerRoundDeadline()
    {
        using var doc = Document("export function main(w) { w.getActiveCell().setValue(42); }");
        doc.Session.Network.Authorize = (_, _) => { Thread.Sleep(5500); return true; };
        doc.EnableScriptsAndCalculate(); doc.RunAction("download"); Assert.Equal("42", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public async Task CancellingPendingHttpDiscardsTheWholeAction()
    {
        var accepted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new ScriptFetchTestServer(async _ => { accepted.TrySetResult(); await release.Task; return new(); });
        using var doc = Document($"export async function main(w) {{ w.getActiveCell().setValue(7); await fetch('{server.Url}'); w.getActiveCell().setValue(42); }}");
        doc.Session.Network.Authorize = (_, _) => true; doc.EnableScriptsAndCalculate();
        using var cancellation = new CancellationTokenSource();
        var action = Task.Run(() => doc.RunAction("download", cancellation.Token));
        try
        {
            await accepted.Task.WaitAsync(TimeSpan.FromSeconds(5)); cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => action);
            Assert.Equal("", doc.GetCellValue(0, 0, 0));
        }
        finally { release.TrySetResult(); }
    }

    [Fact]
    public async Task TypeScriptFetchDeclarationsCompileAndExecuteInTheIsolatedWorker()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(200, "42")));
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.ts", "host.permissions.network.request('127.0.0.1');", ScriptLanguage.TypeScript),
            ["download"] = new("scripts/download.ts", $$"""
                export async function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host): Promise<void> {
                  const request = new Request('{{server.Url}}', { headers: new Headers({ Accept: 'text/plain' }) });
                  const response: Response = await fetch(request);
                  const value: string = await response.text();
                  workbook.getActiveCell().setValue(Number(value));
                }
                """, ScriptLanguage.TypeScript)
        }, ["main", "download"], scriptApiVersion: 4, bindings: [new("main", "initialization"), new("download", "actions")]);
        using var doc = new WorkbookDocument(Book(Sheet("Data")).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.Session.Network.Authorize = (_, _) => true; doc.EnableScriptsAndCalculate(); doc.RunAction("download");
        Assert.Equal("42", doc.GetCellValue(0, 0, 0));
    }

    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public async Task ConfiguredBodyBudgetTransfersLargeBodiesInBothDirections(int bodyMiB)
    {
        var size = bodyMiB * 1024 * 1024;
        await using var server = new ScriptFetchTestServer(request => Task.FromResult(request.Method == "POST"
            ? new ScriptFetchTestServer.Response(200, request.Body.Length.ToString())
            : new ScriptFetchTestServer.Response(200, Bytes: Enumerable.Repeat((byte)251, size).ToArray())));
        using var download = Document($$"""
            export async function main(w) {
              const bytes = new Uint8Array(await (await fetch('{{server.Url}}')).arrayBuffer());
              if(bytes.length !== {{size}} || bytes[0] !== 251 || bytes[bytes.length-1] !== 251) throw Error('large body');
              w.getActiveCell().setValue(bytes.length);
            }
            """);
        if (bodyMiB != 8) download.Session.Network.Configure(new(BodyLimitMiB: bodyMiB));
        download.Session.Network.Authorize = (_, _) => true; download.EnableScriptsAndCalculate(); download.RunAction("download");
        Assert.Equal(size.ToString(), download.GetCellValue(0, 0, 0));
        using var upload = Document($$"""
            export async function main(w) {
              const response = await fetch('{{server.Url}}', {method:'POST',body:new Uint8Array({{size}}).fill(251)});
              w.getActiveCell().setValue(await response.text());
            }
            """);
        if (bodyMiB != 8) upload.Session.Network.Configure(new(BodyLimitMiB: bodyMiB));
        upload.Session.Network.Authorize = (_, _) => true; upload.EnableScriptsAndCalculate(); upload.RunAction("download");
        Assert.Equal(size.ToString(), upload.GetCellValue(0, 0, 0));
    }

    [Fact]
    public async Task ConfiguredActionDeadlineAllowsNetworkWaitBeyondFiveSeconds()
    {
        await using var server = new ScriptFetchTestServer(async _ => { await Task.Delay(5500); return new(200, "slow response"); });
        using var doc = Document($"export async function main(w) {{ w.getActiveCell().setValue(await (await fetch('{server.Url}')).text()); }}");
        doc.Session.Network.Authorize = (_, _) => true; doc.EnableScriptsAndCalculate(); doc.RunAction("download");
        Assert.Equal("slow response", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public async Task CustomDeadlineCancelsSlowHttpWithoutPartialCommit()
    {
        await using var server = new ScriptFetchTestServer(async _ => { await Task.Delay(2500); return new(); });
        using var doc = Document($"export async function main(w) {{ w.getActiveCell().setValue(7); await fetch('{server.Url}'); }}");
        doc.Session.Network.Configure(new(ActionTimeoutSeconds: 1));
        doc.Session.Network.Authorize = (_, _) => true; doc.EnableScriptsAndCalculate();
        Assert.ThrowsAny<Exception>(() => doc.RunAction("download")); Assert.Equal("", doc.GetCellValue(0, 0, 0));
        Assert.Single(server.Requests);
    }

    [Fact]
    public async Task DisabledNetworkRejectsAuthorizationWithoutPromptOrRequest()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response()));
        using var doc = Document($"export async function main() {{ await fetch('{server.Url}'); }}");
        doc.Session.Network.Configure(new(AllowNetwork: false));
        var prompts = 0; doc.Session.Network.Authorize = (_, _) => { prompts++; return true; };
        Assert.ThrowsAny<Exception>(() => doc.EnableScriptsAndCalculate());
        Assert.Equal(0, prompts); Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task TurningOffNetworkCancelsInFlightHttpAndRevokesConsent()
    {
        var accepted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new ScriptFetchTestServer(async _ => { accepted.TrySetResult(); await release.Task; return new(); });
        using var doc = Document($"export async function main(w) {{ w.getActiveCell().setValue(7); await fetch('{server.Url}'); }}");
        var prompts = 0; doc.Session.Network.Authorize = (_, _) => { prompts++; return true; }; doc.EnableScriptsAndCalculate();
        var action = Task.Run(() => doc.RunAction("download"));
        try
        {
            await accepted.Task.WaitAsync(TimeSpan.FromSeconds(5));
            doc.Session.Network.Configure(doc.Session.Network.Options with { AllowNetwork = false });
            await Assert.ThrowsAnyAsync<Exception>(() => action.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.Equal("", doc.GetCellValue(0, 0, 0));
            Assert.ThrowsAny<Exception>(() => doc.EnableScriptsAndCalculate()); Assert.Equal(1, prompts);
        }
        finally { release.TrySetResult(); }
    }
}
