using System.Text;
using System.Text.Json;
using QuickExcel.Scripting;

namespace QuickExcel.Core.Test;

public sealed class ScriptFetchRuntimeTests
{
    private static JscSession Runtime(string source, string initialization = "host.permissions.network.request('example.com');") => new(
        new Dictionary<string, string> { ["main"] = initialization, ["action"] = source }, ["main", "action"], actions: true,
        roles: new Dictionary<string, int> { ["main"] = 4, ["action"] = 2 }, automatic: true, initializationEntries: ["main"]);

    private static JsonElement Request(bool network = false, int maxNetworkBodyBytes = 8 * 1024 * 1024) => JsonSerializer.SerializeToElement(new
    {
        operationId = Guid.NewGuid(), snapshot = new { baseRevision = 0 }, allowNetwork = network, maxNetworkBodyBytes
    });

    private static void Initialize(JscSession runtime)
    {
        runtime.RunAction("@initialize", Request()); runtime.BeginRound();
    }

    [Fact]
    public void FetchPromiseQueuesOutsideNativeTurnAndRetainsCorrectOrdering()
    {
        using var runtime = Runtime("""
            export async function main() {
              const headers = new Headers([['X-Test','one']]); headers.append('x-test','two');
              const request = new Request('https://example.com/data', {method:'POST',headers,body:new URLSearchParams({q:'中文 a'})});
              const promise = fetch(request); console.log('queued', request.bodyUsed);
              const response = await promise;
              const clone = response.clone();
              console.log(response.status, response.ok, response.headers.get('CONTENT-TYPE'));
              console.log(await response.json(), await clone.text(), response.bodyUsed);
              try { await response.text(); throw Error('consumed twice'); } catch(e) { if (!(e instanceof TypeError)) throw e; }
            }
            """);
        Initialize(runtime);
        var requests = 0;
        runtime.HostRequest = json =>
        {
            requests++;
            Assert.Contains("queued true", runtime.GetLogs()); // Called with the VM idle; no reentrant native turn.
            using var parsed = JsonDocument.Parse(json);
            Assert.Equal("network.fetch", parsed.RootElement.GetProperty("method").GetString());
            var args = parsed.RootElement.GetProperty("args")[0];
            Assert.Equal("POST", args.GetProperty("method").GetString());
            Assert.Equal("q=%E4%B8%AD%E6%96%87+a", Encoding.UTF8.GetString(Convert.FromBase64String(args.GetProperty("bodyBase64").GetString()!)));
            return JsonSerializer.Serialize(new { ok = true, value = new { url = "https://example.com/data", status = 404,
                statusText = "Not Found", headers = new[] { new[] { "content-type", "application/json" } }, bodyBase64 = Convert.ToBase64String("{\"value\":42}"u8.ToArray()), redirected = false } });
        };
        runtime.RunAction("action", Request(true));
        Assert.Equal(1, requests);
        Assert.Contains("404 false application/json", runtime.GetLogs());
        Assert.Contains("{value: 42} {\"value\":42} true", runtime.GetLogs());
    }

    [Fact]
    public void AbortBeforeDispatchDoesNotCallHostAndBodyReadersLock()
    {
        using var runtime = Runtime("""
            export async function main() {
              const controller = new AbortController();
              const promise = fetch('https://example.com/', {signal:controller.signal}); controller.abort();
              try { await promise; throw Error('not aborted'); } catch(e) { if(e.name!=='AbortError') throw e; }
              const response = new Response('中文'); const reader = response.body.getReader();
              if(!response.body.locked) throw Error('not locked');
              try { response.clone(); throw Error('clone locked'); } catch(e) { if(!(e instanceof TypeError)) throw e; }
              if(new TextDecoder().decode((await reader.read()).value)!=='中文' || !(await reader.read()).done) throw Error('reader');
              reader.releaseLock(); if(response.body.locked || !response.bodyUsed) throw Error('state');
              console.log('ok');
            }
            """);
        Initialize(runtime); runtime.HostRequest = _ => throw new InvalidOperationException("No HTTP expected.");
        runtime.RunAction("action", Request(true)); Assert.Equal(["ok"], runtime.GetLogs());
    }

    [Fact]
    public void OldResponseBodiesAreEvictedAtAggregateMemoryLimit()
    {
        using var runtime = Runtime("""
            export async function main() {
              const bytes = new Uint8Array(2*1024*1024);
              const responses = Array.from({length:5}, () => new Response(bytes));
              try { await responses[0].arrayBuffer(); throw Error('old body retained'); } catch(e) { if(!(e instanceof TypeError)) throw e; }
              if((await responses[4].arrayBuffer()).byteLength!==bytes.length) throw Error('new body lost');
              console.log('evicted');
            }
            """);
        Initialize(runtime); runtime.RunAction("action", Request(maxNetworkBodyBytes: 2 * 1024 * 1024)); Assert.Equal(["evicted"], runtime.GetLogs());
    }

    [Fact]
    public void PartlyReadBodiesRemainBudgetedAndReleasedReadersResume()
    {
        using var runtime = Runtime("""
            export async function main() {
              const bytes = new Uint8Array(2*1024*1024);
              const old = new Response(bytes), reader = old.body.getReader();
              if((await reader.read()).value.length !== 65536) throw Error('chunk');
              reader.releaseLock(); const resumed = old.body.getReader();
              if((await resumed.read()).value.length !== 65536) throw Error('resume');
              const rest = Array.from({length:4},()=>new Response(bytes));
              try { await resumed.read(); throw Error('reader bypassed eviction'); } catch(e) { if(!(e instanceof TypeError)) throw e; }
              console.log('budgeted');
            }
            """);
        Initialize(runtime); runtime.RunAction("action", Request(maxNetworkBodyBytes: 2 * 1024 * 1024)); Assert.Equal(["budgeted"], runtime.GetLogs());
    }

    [Fact]
    public void ReturnedPromiseAndAsyncInitializationWaitUntilSettled()
    {
        using var runtime = Runtime("export function main(){ return Promise.resolve().then(()=>console.log('action')); }",
            "export async function main(){ await Promise.resolve(); console.log('initialized'); }");
        runtime.RunAction("@initialize", Request()); Assert.Equal(["initialized"], runtime.GetLogs());
        runtime.BeginRound(); runtime.RunAction("action", Request()); Assert.Equal(["action"], runtime.GetLogs());
    }

    [Fact]
    public void RequestsWithoutBodiesRemainReusable()
    {
        using var runtime = Runtime("""
            export async function main() {
              const request = new Request('https://example.com/');
              if(await request.text() !== '' || request.bodyUsed) throw Error('null body consumed');
              await fetch(request); await fetch(request);
              if(request.bodyUsed) throw Error('null body sent');
            }
            """);
        Initialize(runtime); var calls = 0;
        runtime.HostRequest = _ => { calls++; return "{\"ok\":true,\"value\":{\"url\":\"https://example.com/\",\"status\":204,\"statusText\":\"No Content\",\"headers\":[],\"bodyBase64\":\"\",\"redirected\":false}}"; };
        runtime.RunAction("action", Request(true)); Assert.Equal(2, calls);
    }

    [Fact]
    public void EightMiBTextUploadFitsNativeTurnAndUsesConfiguredBodyLimit()
    {
        using var runtime = Runtime("""
            export async function main() {
              const body = '中'.repeat(2*1024*1024) + 'a'.repeat(2*1024*1024);
              await fetch('https://example.com/upload', {method:'POST',body});
              console.log('uploaded');
            }
            """);
        Initialize(runtime); var calls = 0;
        runtime.HostRequest = json =>
        {
            calls++;
            using var call = JsonDocument.Parse(json);
            var bytes = Convert.FromBase64String(call.RootElement.GetProperty("args")[0].GetProperty("bodyBase64").GetString()!);
            Assert.Equal(8 * 1024 * 1024, bytes.Length);
            Assert.Equal("中", Encoding.UTF8.GetString(bytes, 0, 3)); Assert.Equal((byte)'a', bytes[^1]);
            return "{\"ok\":true,\"value\":{\"url\":\"https://example.com/upload\",\"status\":204,\"statusText\":\"No Content\",\"headers\":[],\"bodyBase64\":\"\",\"redirected\":false}}";
        };
        runtime.RunAction("action", Request(true)); Assert.Equal(1, calls); Assert.Equal(["uploaded"], runtime.GetLogs());
        Assert.False(runtime.LastOperationInvalidated);
    }

    [Fact]
    public void EightMiBTextResponseDecodesWithinNativeTurn()
    {
        using var runtime = Runtime("""
            export async function main() {
              const response = await fetch('https://example.com/download');
              const text = await response.text();
              if(text.length !== 4*1024*1024 || text[0] !== '中' || text[text.length-1] !== 'a') throw Error('decoded body');
              console.log('downloaded');
            }
            """);
        Initialize(runtime);
        var bytes = Encoding.UTF8.GetBytes(new string('中', 2 * 1024 * 1024) + new string('a', 2 * 1024 * 1024));
        Assert.Equal(8 * 1024 * 1024, bytes.Length);
        runtime.HostRequest = _ => JsonSerializer.Serialize(new { ok = true, value = new { url = "https://example.com/download", status = 200,
            statusText = "OK", headers = Array.Empty<string[]>(), bodyBase64 = Convert.ToBase64String(bytes), redirected = false } });
        runtime.RunAction("action", Request(true)); Assert.Equal(["downloaded"], runtime.GetLogs());
        Assert.False(runtime.LastOperationInvalidated);
    }

    [Theory]
    [InlineData(true)] [InlineData(false)]
    public void EightMiBBinaryBodiesAvoidJsonEscapingExpansion(bool upload)
    {
        using var runtime = Runtime(upload ? """
            export async function main() {
              await fetch('https://example.com/upload', {method:'POST',body:new Uint8Array(8*1024*1024).fill(251)});
              console.log('binary uploaded');
            }
            """ : """
            export async function main() {
              const response = await fetch('https://example.com/download');
              const bytes = new Uint8Array(await response.arrayBuffer());
              if(bytes.length !== 8*1024*1024 || bytes[0] !== 251 || bytes[bytes.length-1] !== 251) throw Error('binary body');
              console.log('binary downloaded');
            }
            """);
        Initialize(runtime);
        runtime.HostRequest = json =>
        {
            if (upload)
            {
                Assert.True(Encoding.UTF8.GetByteCount(json) < 16 * 1024 * 1024);
                using var call = JsonDocument.Parse(json);
                var encoded = call.RootElement.GetProperty("args")[0].GetProperty("bodyBase64").GetString()!;
                Assert.DoesNotContain('+', encoded); Assert.Contains('-', encoded);
                var bytes = Convert.FromBase64String(encoded.Replace('-', '+').Replace('_', '/'));
                Assert.Equal(8 * 1024 * 1024, bytes.Length); Assert.Equal(251, bytes[0]); Assert.Equal(251, bytes[^1]);
                return "{\"ok\":true,\"value\":{\"url\":\"https://example.com/upload\",\"status\":204,\"statusText\":\"No Content\",\"headers\":[],\"bodyBase64\":\"\",\"redirected\":false}}";
            }
            return JsonSerializer.Serialize(new { ok = true, value = new { url = "https://example.com/download", status = 200,
                statusText = "OK", headers = Array.Empty<string[]>(),
                bodyBase64 = Convert.ToBase64String(Enumerable.Repeat((byte)251, 8 * 1024 * 1024).ToArray()).Replace('+', '-').Replace('/', '_'), redirected = false } });
        };
        runtime.RunAction("action", Request(true)); Assert.Equal([upload ? "binary uploaded" : "binary downloaded"], runtime.GetLogs());
        Assert.False(runtime.LastOperationInvalidated);
    }

    [Fact]
    public void HostLimitUpdatesEachActionAndCannotBeRaisedByScriptGlobals()
    {
        using var runtime = Runtime("""
            let round = 0;
            export async function main() {
              globalThis.maxFetchBody = 100*1024*1024;
              const body = new Uint8Array(1024*1024+1);
              let accepted = false;
              try { new Response(body); accepted = true; } catch(e) { if(!(e instanceof TypeError)) throw e; }
              if(accepted !== (++round === 2)) throw Error('host limit ignored');
            }
            """);
        Initialize(runtime);
        runtime.RunAction("action", Request(maxNetworkBodyBytes: 1024 * 1024));
        runtime.BeginRound(); runtime.RunAction("action", Request(maxNetworkBodyBytes: 8 * 1024 * 1024));
    }

    [Theory]
    [InlineData(0)] [InlineData(1024 * 1024 - 1)] [InlineData(10 * 1024 * 1024 + 1)]
    public void InvalidHostBodyLimitIsRejectedBeforeExecutingAction(int limit)
    {
        using var runtime = Runtime("export function main(){console.log('must not run');}"); Initialize(runtime);
        Assert.Throws<InvalidOperationException>(() => runtime.RunAction("action", Request(maxNetworkBodyBytes: limit)));
        Assert.Empty(runtime.GetLogs());
    }

    [Theory]
    [InlineData("return Promise.reject(Error('rejected'));", "rejected")]
    [InlineData("return new Promise(()=>{});", "pending")]
    [InlineData("return Promise.resolve().then(()=>{while(true){}});", "time limit")]
    public void RejectedOrUnsettleablePromisesFailWithinLimits(string body, string message)
    {
        using var runtime = Runtime("export function main(){" + body + "}"); Initialize(runtime);
        Assert.Contains(message, Assert.Throws<InvalidOperationException>(() => runtime.RunAction("action", Request())).Message);
    }

    [Fact]
    public void ImportedHelperCannotDeclareInitializationRootPermissions()
    {
        var error = Assert.Throws<InvalidOperationException>(() => new JscSession(
            new Dictionary<string, string> { ["main"] = "import 'helper';", ["helper"] = "host.permissions.network.request('example.com');" },
            ["main"], actions: true, roles: new Dictionary<string, int> { ["main"] = 4, ["helper"] = 0 },
            automatic: true, initializationEntries: ["main"]));
        Assert.Contains("initialization root", error.Message);
    }
}
