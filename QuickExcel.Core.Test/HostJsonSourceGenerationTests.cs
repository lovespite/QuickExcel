using System.Collections.Immutable;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using QuickExcel.Domain;
using QuickExcel.Presentation.Web;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Security;
using QuickExcel.Serialization;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class HostJsonSourceGenerationTests
{
    [Fact]
    public void JsonNodeManifestRoundTripPreservesStorageAndVirtualFileFeatures()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "json-node-evidence", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var io = new TypedWorkbookStorage();
            var files = VirtualFileSet.Empty.Apply([new("write", "//xlsq/data/note.txt", Encoding.UTF8.GetBytes("内容").ToImmutableArray())]);
            var snapshot = Book(Sheet("S")).WithPackageFiles(files).WithLocalStorage(ScriptStorage.Empty.Set("中文", "值"));
            var captured = io.Capture(snapshot);
            var manifest = JsonNode.Parse(captured.Entries["//xlsq/manifest.json"].Data.ToArray())!.AsObject();
            manifest["optionalMarker"] = "保留";
            var edited = io.Apply(snapshot, [new("write", "//xlsq/manifest.json", Encoding.UTF8.GetBytes(manifest.ToJsonString()).ToImmutableArray())]);
            var path = Path.Combine(directory, "roundtrip.xlsq");
            io.Save(path, edited);
            var reopened = io.Open(path);
            Assert.Equal("值", reopened.LocalStorage.Get("中文"));
            Assert.Equal(files.Entries["//xlsq/data/note.txt"].Data.ToArray(), reopened.PackageFiles.Entries["//xlsq/data/note.txt"].Data.ToArray());
            using var zip = System.IO.Compression.ZipFile.OpenRead(path);
            using var stored = JsonDocument.Parse(zip.GetEntry("manifest.json")!.Open());
            var features = stored.RootElement.GetProperty("requiredFeatures").EnumerateArray().Select(p => p.GetString()).ToArray();
            Assert.Contains("script-storage", features);
            Assert.Contains("virtual-files", features);
            Assert.Equal("保留", stored.RootElement.GetProperty("optionalMarker").GetString());
            Assert.Single(stored.RootElement.GetProperty("localStorage").EnumerateArray());
            Assert.Contains(stored.RootElement.GetProperty("files").EnumerateArray(), p => p.GetProperty("path").GetString() == "data/note.txt");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void GeneratedCacheAndConsentMetadataPreservesAuthenticatedBytes()
    {
        var policy = new ScriptCertificateTrustStore.ExecutionPolicyIdentity(WorkerProtocol.Engine,
            WorkerProtocol.Policy, WorkerProtocol.Version, CompilerProtocol.Identity, CompilerProtocol.Policy, CompilerProtocol.Options);
        var originalPolicy = JsonSerializer.Serialize(new
        {
            engine = policy.engine, policy = policy.policy, protocol = policy.protocol,
            compiler = policy.compiler, compilerPolicy = policy.compilerPolicy, options = policy.options
        });
        Assert.Equal(originalPolicy, JsonSerializer.Serialize(policy, InfrastructureJsonContext.Default.ExecutionPolicyIdentity));

        var compilation = new ScriptCompilationCache.CompilationMetadata("编译器", "source", "output", "compact",
            "typescript", "checkEmit", 1.25, [new("模块", 0, 123, 45)]);
        var authentication = new ScriptCompilationCache.Authentication("QuickExcel.ScriptCompilationCache", 2,
            "signature", compilation, 168, "data", "publisher");
        var legacyOptions = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            MaxDepth = 24, PropertyNameCaseInsensitive = false, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
            RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true
        };
        var originalCache = JsonSerializer.SerializeToUtf8Bytes(new
        {
            purpose = authentication.Purpose, authentication.Version, authentication.SignatureSha256,
            authentication.Compilation, authentication.DataBytes, authentication.DataSha256, authentication.Protection
        }, legacyOptions);
        Assert.Equal(originalCache, JsonSerializer.SerializeToUtf8Bytes(authentication, InfrastructureJsonContext.Protocol.Authentication));
        var signature = new ScriptSignature(2, "RSA-SHA256", "certificate", "value");
        Assert.Equal(JsonSerializer.SerializeToUtf8Bytes(signature, legacyOptions),
            JsonSerializer.SerializeToUtf8Bytes(signature, InfrastructureJsonContext.Protocol.ScriptSignature));
    }

    [Fact]
    public void GeneratedProtectedPreferencesRetainNullOmissionAndStrictFields()
    {
        var envelope = new ScriptCertificateTrustStore.ProtectedTrust(1, [1, 2, 3]);
        Assert.Equal("{\"Version\":1,\"Data\":\"AQID\"}",
            JsonSerializer.Serialize(envelope, InfrastructureJsonContext.Strict.ProtectedTrust));
        Assert.Throws<JsonException>(() => JsonSerializer.Deserialize(
            "{\"Version\":1,\"Data\":\"AQID\",\"unexpected\":true}", InfrastructureJsonContext.Strict.ProtectedTrust));
        var preferences = new ScriptSigningPreferences(null, null, null);
        Assert.Equal(JsonSerializer.Serialize(preferences),
            JsonSerializer.Serialize(preferences, InfrastructureJsonContext.Strict.ScriptSigningPreferences));
    }

    [Theory]
    [InlineData(3, false, false)]
    [InlineData(3, false, true)]
    [InlineData(3, true, true)]
    [InlineData(4, false, true)]
    [InlineData(4, true, true)]
    public void StorageScriptDigestPreservesLegacyPropertyOrderAndBindingCasing(int api, bool typescript, bool actions)
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main." + (typescript ? "ts" : "js"), "// 中文 <>&\nexport function RUN() {}",
                typescript ? ScriptLanguage.TypeScript : ScriptLanguage.JavaScript),
            ["helper"] = new("scripts/helper.js", "export const text = '辅助';")
        }, ["main"], scriptApiVersion: api, bindings: actions ? [new("main", "actions")] : null);
        var capture = new TypedWorkbookStorage().Capture(Book(Sheet("S")).WithScriptModules(modules));
        using var state = JsonDocument.Parse(capture.Entries["//xlsq/state/document.json"].Data.ToArray());
        Assert.Equal(LegacyScriptHash(modules), state.RootElement.GetProperty("calculation").GetProperty("scriptSetSha256").GetString());
    }

    [Fact]
    public void WebSourceGenerationRetainsWebNumericStringDefaultsAndDraftFields()
    {
        const string body = "{\"protocol\":\"1\",\"instance\":\"view\",\"document\":\"doc\",\"generation\":\"3\",\"type\":\"draftAck\",\"requestId\":\"r1\",\"payload\":{}}";
        Assert.Equal(3, WebMessage.Parse(body, "view", "doc", 3, new HashSet<string> { "draftAck" }).Generation);
        var chunk = new DraftChunk("main", 1, 2, "transfer", 0, 1, 4, "ABC", "数据");
        var encoded = WebMessage.Create("view", "doc", 3, "draftChunk", "r1", chunk);
        using var document = JsonDocument.Parse(encoded);
        Assert.Equal(JsonSerializer.Serialize(chunk, new JsonSerializerOptions(JsonSerializerDefaults.Web)),
            document.RootElement.GetProperty("payload").GetRawText());
    }

    // Independent pre-source-generation representation, including the historical
    // uppercase TypeScript binding fields. Existing workbook digests must stay valid.
    private static string LegacyScriptHash(ScriptModuleSet modules)
    {
        static string Hash(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));
        if (modules.HasTypeScript) return Hash(JsonSerializer.SerializeToUtf8Bytes(new
        {
            scriptApiVersion = modules.ScriptApiVersion, compilerOptions = ScriptModuleSet.TypeScriptOptions, entries = modules.EntryModules,
            modules = modules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal).Select(m => new
            {
                id = m.Key, path = m.Value.Path, language = m.Value.Language == ScriptLanguage.TypeScript ? "typescript" : "javascript",
                sha256 = Hash(Encoding.UTF8.GetBytes(m.Value.Source))
            }),
            bindings = modules.Bindings.OrderBy(b => b.Module, StringComparer.Ordinal).ThenBy(b => b.Kind, StringComparer.Ordinal)
        }));
        var sources = modules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal).Select(m => new
        { id = m.Key, path = m.Value.Path, sha256 = Hash(Encoding.UTF8.GetBytes(m.Value.Source)) });
        return modules.HasActions
            ? Hash(JsonSerializer.SerializeToUtf8Bytes(new
            {
                scriptApiVersion = modules.ScriptApiVersion, entries = modules.EntryModules, modules = sources,
                bindings = modules.Bindings.OrderBy(b => b.Module, StringComparer.Ordinal).ThenBy(b => b.Kind, StringComparer.Ordinal)
                    .Select(b => new { module = b.Module, kind = b.Kind })
            }))
            : Hash(JsonSerializer.SerializeToUtf8Bytes(new { scriptApiVersion = modules.ScriptApiVersion, entries = modules.EntryModules, modules = sources }));
    }
}
