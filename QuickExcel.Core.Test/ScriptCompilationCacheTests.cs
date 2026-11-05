using System.Collections.Immutable;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Security;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptCompilationCacheTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-compilation-cache-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static CalculationStamp Stamp => new(Guid.NewGuid(), 0, 0, 0);
    public ScriptCompilationCacheTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);

    private static ScriptModuleSet Modules(bool typescript = true, int api = 4) => new(new Dictionary<string, ScriptModule>
    {
        ["main"] = new("scripts/main." + (typescript ? "ts" : "js"), api == 4
            ? "import {value} from 'helper'; export function VALUE() { return value; }"
            : "import {value} from 'helper'; host.functions.register('VALUE', () => value);", typescript ? ScriptLanguage.TypeScript : ScriptLanguage.JavaScript),
        ["helper"] = new("scripts/helper.js", "export const value = 42;")
    }, ["main"], scriptApiVersion: api);

    private async Task<string> SignedFile(bool typescript = true, int api = 4)
    {
        var modules = Modules(typescript, api);
        using var certificate = ScriptSignatures.CreateSelfSigned("Compilation cache tests");
        var signature = await ScriptSignatures.SignAsync(modules, certificate, Stamp);
        Assert.Equal(1, ScriptCompilation.ProcessLaunchCount(modules));
        modules = modules.WithSignature(signature);
        Assert.False(ScriptCompilation.IsCacheDirty(modules));
        var path = Path.Combine(_folder, "signed.xlsq");
        _storage.Save(path, Book(Sheet("S")).WithScriptModules(modules));
        return path;
    }

    [Theory]
    [InlineData(true, 3)]
    [InlineData(true, 4)]
    [InlineData(false, 3)]
    [InlineData(false, 4)]
    public async Task SigningPersistsAuthenticatedArtifactsAndReopenExecutesWithoutCompiler(bool typescript, int api)
    {
        var path = await SignedFile(typescript, api);
        using (var zip = ZipFile.OpenRead(path))
        {
            Assert.NotNull(zip.GetEntry(ScriptCompilationCache.DataPart));
            var metadata = zip.GetEntry(ScriptCompilationCache.Part)!;
            Assert.InRange(metadata.Length, 1, ScriptCompilationCache.MaximumMetadataBytes);
            using var json = JsonDocument.Parse(metadata.Open());
            Assert.Equal(2, json.RootElement.GetProperty("version").GetInt32());
            Assert.False(json.RootElement.GetProperty("compilation").GetProperty("outputs")[0].TryGetProperty("javascript", out _));
        }
        var snapshot = _storage.Open(path);
        var modules = snapshot.ScriptModules;
        Assert.Equal(Modules(typescript, api).Modules.OrderBy(m => m.Key), modules.Modules.OrderBy(m => m.Key));
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, ScriptSignatures.Verify(modules).Status);
        Assert.False(ScriptCompilation.IsCacheDirty(modules));
        using var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, Stamp, TimeSpan.FromSeconds(5), default));
        Assert.Equal(42d, runtime.Invoke(api == 4 ? "JS.main.VALUE" : "VALUE", []));
        Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(modules));
        Assert.DoesNotContain(_storage.Capture(snapshot).Entries.Keys, p => p.Contains("cache/"));
        // A consumer cannot modify the privately stored accepted reply.
        var result = (await ScriptCompilation.PrepareAsync(modules, Stamp, default))!;
        result.Result.Outputs[0] = result.Result.Outputs[0] with { Javascript = "throw new Error('injected');" };
        Assert.DoesNotContain((await ScriptCompilation.PrepareAsync(modules, Stamp, default))!.Result.Outputs, o => o.Javascript.Contains("injected"));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("hash")]
    [InlineData("json")]
    [InlineData("version")]
    [InlineData("compiler")]
    [InlineData("signature")]
    [InlineData("source")]
    [InlineData("proof")]
    [InlineData("forged-output-and-hashes")]
    [InlineData("malformed-map")]
    public async Task InvalidCacheRecompilesAndReplacesBeforeExecutionThenSurvivesReopen(string mutation)
    {
        var path = await SignedFile();
        Mutate(path, mutation);
        var snapshot = _storage.Open(path);
        var modules = snapshot.ScriptModules;
        Assert.True(ScriptCompilation.IsCacheDirty(modules));
        var originalSignature = modules.Signature;
        using (var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, Stamp, TimeSpan.FromSeconds(5), default)))
            Assert.Equal(42d, runtime.Invoke("JS.main.VALUE", []));
        Assert.Equal(1, ScriptCompilation.ProcessLaunchCount(modules));
        Assert.False(ScriptCompilation.IsCacheDirty(modules));
        Assert.Equal(originalSignature, modules.Signature);
        _storage.Save(path, snapshot);
        var reopened = _storage.Open(path).ScriptModules;
        Assert.False(ScriptCompilation.IsCacheDirty(reopened));
        await ScriptCompilation.PrepareAsync(reopened, Stamp, default);
        Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(reopened));
    }

    [Fact]
    public async Task LegacySignatureWithoutArtifactsIsCompiledOnceWithoutResigning()
    {
        var modules = Modules();
        using var certificate = ScriptSignatures.CreateSelfSigned("Legacy source signature");
        modules = modules.WithSignature(ScriptSignatures.Sign(modules, certificate));
        Assert.True(ScriptCompilation.IsCacheDirty(modules));
        await ScriptCompilation.PrepareAsync(modules, Stamp, default);
        await ScriptCompilation.PrepareAsync(modules, Stamp, default);
        Assert.Equal(1, ScriptCompilation.ProcessLaunchCount(modules));
        var path = Path.Combine(_folder, "legacy.xlsq");
        _storage.Save(path, Book(Sheet("S")).WithScriptModules(modules));
        var reopened = _storage.Open(path).ScriptModules;
        Assert.Equal(modules.Signature, reopened.Signature);
        Assert.False(ScriptCompilation.IsCacheDirty(reopened));
        // A copied/forged local receipt cannot authenticate a modified artifact.
        Mutate(path, "forged-output-and-hashes");
        Assert.True(ScriptCompilation.IsCacheDirty(_storage.Open(path).ScriptModules));
    }

    [Theory]
    [InlineData("source")]
    [InlineData("language")]
    [InlineData("binding")]
    [InlineData("virtual-source")]
    public async Task EditingInvalidatesCacheAndSignatureAndDoesNotSerializeOldArtifacts(string change)
    {
        var path = await SignedFile();
        var snapshot = _storage.Open(path);
        using var session = new DocumentSession(snapshot);
        session.SetScriptsEnabled(true);
        var prior = snapshot.ScriptModules;
        if (change == "source") session.Execute(new UpdateScriptSourceCommand("export function VALUE() { return 7; }"));
        else if (change == "virtual-source")
            snapshot = _storage.Apply(snapshot, [new("write", "//xlsq/scripts/main.ts", Encoding.UTF8.GetBytes("export function VALUE() { return 7; }").ToImmutableArray())]);
        else session.Execute(new UpdateScriptModulesCommand(new(prior.Modules.SetItem("main", prior.Modules["main"] with
            { Language = change == "language" ? ScriptLanguage.JavaScript : ScriptLanguage.TypeScript }), prior.EntryModules,
            scriptApiVersion: 4, bindings: [new("main", change == "binding" ? "actions" : "formula-functions")])));
        if (change != "virtual-source") { snapshot = session.Snapshot; Assert.False(session.ScriptsEnabled); }
        Assert.Null(snapshot.ScriptModules.Signature);
        Assert.True(ScriptCompilation.IsCacheDirty(snapshot.ScriptModules));
        Assert.False(ScriptCompilation.IsCacheDirty(prior));
        _storage.Save(path, snapshot);
        using var zip = ZipFile.OpenRead(path);
        Assert.Null(zip.GetEntry(ScriptCompilationCache.Part));
        Assert.Null(zip.GetEntry(ScriptCompilationCache.DataPart));
    }

    [Fact]
    public async Task InvalidSourcesOrCancellationCannotSignOrExecuteAndOriginalSignatureCannotBeReattached()
    {
        using var certificate = ScriptSignatures.CreateSelfSigned("Failed compilation");
        var bad = Modules().WithPrimarySource("const n: number = 'wrong'; export function VALUE() { return n; }");
        await Assert.ThrowsAsync<ScriptCompilationException>(() => ScriptSignatures.SignAsync(bad, certificate, Stamp));
        Assert.Null(bad.Signature);
        Assert.True(ScriptCompilation.IsCacheDirty(bad));
        var good = Modules();
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ScriptSignatures.SignAsync(good, certificate, Stamp, cancel.Token));
        Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(good));
        var signed = _storage.Open(await SignedFile()).ScriptModules;
        var tampered = signed.WithPrimarySource("export function VALUE() { return 0; }").WithSignature(signed.Signature);
        Assert.True(ScriptCompilation.IsCacheDirty(tampered));
        await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptCompilation.PrepareAsync(tampered, Stamp, default));
        Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(tampered));
    }

    [Fact]
    public async Task UncommittedDraftInvalidationDropsPersistedAndSessionArtifacts()
    {
        var path = await SignedFile();
        var modules = _storage.Open(path).ScriptModules;
        var signature = modules.Signature;
        ScriptCompilation.InvalidateCache(modules);
        Assert.True(ScriptCompilation.IsCacheDirty(modules));
        Assert.Null(ScriptCompilationCache.Export(modules));
        Assert.Equal(signature, modules.Signature); // The immutable committed sources are still the original sources.
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ScriptCompilation.PrepareAsync(modules, Stamp, cancelled.Token));
        Assert.True(ScriptCompilation.IsCacheDirty(modules));
        Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(modules));
    }

    [Fact]
    public async Task SigningDoesNotExecuteInitializationAndCacheCannotBeWrittenThroughVirtualFiles()
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        { ["main"] = new("scripts/main.ts", "throw new Error('must not run while signing'); export function main() {}", ScriptLanguage.TypeScript) },
            ["main"], scriptApiVersion: 4, bindings: [new("main", "initialization")]);
        using var certificate = ScriptSignatures.CreateSelfSigned("No execution while signing");
        modules = modules.WithSignature(await ScriptSignatures.SignAsync(modules, certificate, Stamp));
        Assert.False(ScriptCompilation.IsCacheDirty(modules));
        var snapshot = Book(Sheet("S")).WithScriptModules(modules);
        Assert.Throws<ArgumentException>(() => _storage.Apply(snapshot, [new("write", "//xlsq/" + ScriptCompilationCache.Part, [1, 2, 3])]));
        Assert.Throws<ArgumentException>(() => _storage.Apply(snapshot, [new("write", "//xlsq/" + ScriptCompilationCache.DataPart, [1, 2, 3])]));
        Assert.Throws<ArgumentException>(() => _storage.Apply(snapshot, [new("mkdir", "//xlsq/" + ScriptCompilationCache.Part)]));
        Assert.Throws<ArgumentException>(() => _storage.Apply(snapshot, [new("write", "//xlsq/cache", [1])]));
        // Ordinary virtual cache directories/files are user content and remain supported.
        snapshot = _storage.Apply(snapshot, [new("mkdir", "//xlsq/cache"), new("write", "//xlsq/cache/user.txt", [42])]);
        var path = Path.Combine(_folder, "virtual-cache.xlsq");
        _storage.Save(path, snapshot);
        var reopened = _storage.Open(path);
        Assert.True(snapshot.PackageFiles.ContentEquals(reopened.PackageFiles));
        Assert.False(ScriptCompilation.IsCacheDirty(reopened.ScriptModules));
    }

    [Fact]
    public async Task ActionsAndUdfsShareArtifactsButRetainSeparateExecutionContexts()
    {
        var modules = Modules();
        modules = new(modules.Modules.Add("action", new("scripts/action.ts", "export function main(workbook: QuickExcel.Workbook) { workbook.getActiveWorksheet().getRange('A1').setValue(42); }", ScriptLanguage.TypeScript)),
            ["main", "action"], scriptApiVersion: 4, bindings: [new("main", "formula-functions"), new("action", "actions")]);
        using var certificate = ScriptSignatures.CreateSelfSigned("Shared compilation");
        modules = modules.WithSignature(await ScriptSignatures.SignAsync(modules, certificate, Stamp));
        var path = Path.Combine(_folder, "actions.xlsq");
        _storage.Save(path, Book(Sheet("S")).WithScriptModules(modules));
        var snapshot = _storage.Open(path);
        using var document = new WorkbookDocument(snapshot, new(_storage, new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        document.EnableScriptsAndCalculate();
        document.RunAction("action");
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
        Assert.Equal(0, ScriptCompilation.ProcessLaunchCount(snapshot.ScriptModules));
    }

    private static void Mutate(string path, string mutation)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        JsonNode Read(string name) { using var reader = new StreamReader(zip.GetEntry(name)!.Open()); return JsonNode.Parse(reader.ReadToEnd())!; }
        void Write(string name, byte[] bytes) { zip.GetEntry(name)?.Delete(); using var output = zip.CreateEntry(name).Open(); output.Write(bytes); }
        var manifest = Read("manifest.json");
        var cache = Read(ScriptCompilationCache.Part);
        if (mutation == "missing") { zip.GetEntry(ScriptCompilationCache.DataPart)!.Delete(); return; }
        if (mutation == "version") cache["version"] = 999;
        if (mutation == "compiler") cache["compilation"]!["compiler"] = new string('0', 64);
        if (mutation == "signature") cache["signatureSha256"] = new string('0', 64);
        if (mutation == "source") cache["compilation"]!["sourceSha256"] = new string('0', 64);
        if (mutation == "proof") cache["proof"] = Convert.ToBase64String(new byte[384]);
        if (mutation is "forged-output-and-hashes" or "malformed-map")
        {
            byte[] data;
            using (var input = zip.GetEntry(ScriptCompilationCache.DataPart)!.Open())
            using (var buffer = new MemoryStream()) { input.CopyTo(buffer); data = buffer.ToArray(); }
            var compilation = cache["compilation"]!;
            var outputs = compilation["outputs"]!.AsArray().Select(o =>
            {
                var offset = o!["offset"]!.GetValue<int>(); var jsBytes = o["javascriptBytes"]!.GetValue<int>();
                return new CompilerOutput(o["id"]!.GetValue<string>(),
                    Encoding.UTF8.GetString(data, offset, jsBytes), Encoding.UTF8.GetString(data, offset + jsBytes, o["sourceMapBytes"]!.GetValue<int>()));
            }).Select(o => mutation == "malformed-map" ? o with { SourceMap = "null" }
                : o with { Javascript = o.Javascript.Replace("42", "999") }).ToArray();
            var indices = new JsonArray();
            using var updated = new MemoryStream();
            foreach (var output in outputs)
            {
                var js = Encoding.UTF8.GetBytes(output.Javascript); var map = Encoding.UTF8.GetBytes(output.SourceMap);
                indices.Add(new JsonObject { ["id"] = output.Id, ["offset"] = (int)updated.Position, ["javascriptBytes"] = js.Length, ["sourceMapBytes"] = map.Length });
                updated.Write(js); updated.Write(map);
            }
            data = updated.ToArray();
            compilation["outputs"] = indices;
            compilation["outputSha256"] = CompilerProtocol.OutputDigest(new("5.9.3", "check", true, 0, [], outputs));
            cache["dataBytes"] = data.Length;
            cache["dataSha256"] = Convert.ToHexString(SHA256.HashData(data));
            Write(ScriptCompilationCache.DataPart, data);
            manifest["scriptCompilationData"]!["sha256"] = Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
        }
        var bytes = mutation == "json" ? Encoding.UTF8.GetBytes("{broken") : JsonSerializer.SerializeToUtf8Bytes(cache);
        Write(ScriptCompilationCache.Part, bytes);
        manifest["scriptCompilation"]!["sha256"] = mutation == "hash" ? new string('0', 64) : Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        Write("manifest.json", JsonSerializer.SerializeToUtf8Bytes(manifest));
    }
}
