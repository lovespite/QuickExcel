using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Core.Test;

public sealed class WorkerJsonSourceGenerationTests
{
    // This deliberately uses the old reflection encoder as an independent
    // compatibility oracle. Production protocol code has no reflection fallback.
    private static readonly JsonSerializerOptions Legacy = new(JsonSerializerDefaults.Web)
    {
        MaxDepth = 24,
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true
    };

    [Fact]
    public void ModuleDigestsRetainLegacyFieldOrderEscapingAndOrdinalSort()
    {
        var bundle = new LoadModules([
            new("z", "scripts/中文.js", "export const x = '<&🧮>';", "中文 <&"),
            new("A", "scripts/a.js", "export const y = 2;", "")], ["z", "A"], "")
        {
            ApiVersion = 4, FunctionEntries = ["A"], InitializationEntries = ["z"]
        };
        var oldBytes = JsonSerializer.SerializeToUtf8Bytes(new
        {
            apiVersion = bundle.ApiVersion, entries = bundle.EntryModules,
            functionEntries = bundle.FunctionEntries, actionEntries = bundle.ActionEntries,
            eventEntries = bundle.EventEntries, initializationEntries = bundle.InitializationEntries,
            modules = bundle.Modules.OrderBy(m => m.Id, StringComparer.Ordinal).Select(m => new
            {
                id = m.Id, path = m.Path,
                sha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(m.Source))), description = m.Description
            })
        });
        Assert.Equal(Convert.ToHexString(SHA256.HashData(oldBytes)), bundle.ComputeDigest());
        (bundle with { Sha256 = bundle.ComputeDigest() }).Validate();
    }

    [Theory]
    [InlineData("none")]
    [InlineData("compact")]
    [InlineData("full")]
    public void CompilerSourceAndOutputDigestsRetainLegacyBytes(string sourceMapMode)
    {
        var request = new CompileRequest([
            new("z", "typescript", "const x: string = '中文 <&🧮>';"),
            new("A", "javascript", "export const a = 1;")], ["z", "A"], "")
        { ApiVersion = 4, SourceMapMode = sourceMapMode };
        var oldJson = JsonSerializer.Serialize(new
        {
            compiler = CompilerProtocol.Identity, compilerPolicy = CompilerProtocol.Policy, options = CompilerProtocol.Options,
            engine = WorkerProtocol.Engine, policy = WorkerProtocol.Policy, api = request.ApiVersion, sourceMapMode = request.SourceMapMode,
            modules = request.Modules.OrderBy(m => m.Id, StringComparer.Ordinal), entries = request.Entries
        }, Legacy);
        Assert.Equal(CompilerProtocol.Hash(oldJson), request.ComputeDigest());
        (request with { Sha256 = request.ComputeDigest() }).Validate();
        var result = new CompilerResult("5.9.3", "check", true, 1.5, [],
            [new("z", "const x = '中文 <&🧮>';", "{\"mappings\":\"AAAA\"}")]);
        Assert.Equal(CompilerProtocol.Hash(JsonSerializer.Serialize(result.Outputs, Legacy)), CompilerProtocol.OutputDigest(result));
    }

    [Fact]
    public async Task GeneratedEnvelopeMatchesLegacyBytesAndRetainsNullableMetadata()
    {
        var identity = new RequestIdentity(Guid.Parse("a05de35b-32bf-4c9e-87cd-bc19389a2ce1"),
            Guid.Parse("425ae889-f647-4948-8f8e-f6f1008d281b"),
            Guid.Parse("96ffbdc2-55a2-4788-bd7c-7cc9c390b268"), 3, null, 4)
        { RoundId = Guid.Parse("6b95d5dc-d2a1-48ba-8447-92b3893e63f6") };
        var payload = new WorkerFailure("Action", "中文 <&🧮") { Logs = ["first", ""] };
        var message = WorkerProtocol.Message(identity, "failure", payload);
        Assert.Equal(JsonSerializer.Serialize(payload, Legacy), message.Payload.GetRawText());
        using var stream = new MemoryStream();
        await WorkerProtocol.WriteAsync(stream, message);
        Assert.Equal(JsonSerializer.SerializeToUtf8Bytes(message, Legacy), stream.ToArray()[4..]);
        stream.Position = 0;
        var decoded = await WorkerProtocol.ReadAsync(stream);
        Assert.Equal(identity, decoded.Identity);
        Assert.Equal(payload.Logs, WorkerProtocol.Payload<WorkerFailure>(decoded).Logs);
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"source\":null,\"sha256\":\"hash\"}")]
    [InlineData("{\"Source\":\"x\",\"sha256\":\"hash\"}")]
    [InlineData("{\"source\":\"x\",\"sha256\":\"hash\",\"extra\":true}")]
    public void GeneratedPayloadKeepsRequiredNullabilityCasingAndUnknownFieldChecks(string json)
    {
        using var document = JsonDocument.Parse(json);
        var message = new WorkerEnvelope(WorkerProtocol.Version, M4ProtocolTests.Identity(), "load", document.RootElement);
        Assert.Throws<ProtocolException>(() => WorkerProtocol.Payload<LoadScript>(message));
    }

    [Fact]
    public void PrimitiveContainersAndExplicitRangeHaveNoReflectionDependency()
    {
        object?[] values = [null, "中文", 1d, true, 2, 3L, 4m];
        var array = WireValue.FromObject(values);
        Assert.Equal("array", array.Kind);
        Assert.Equal(JsonSerializer.Serialize(values, Legacy), array.Value.GetRawText());
        var range = new WorkerRangeValue("range", 1, values.Length, values);
        var wire = WireValue.FromObject(range);
        Assert.Equal("range", wire.Kind);
        Assert.Equal(JsonSerializer.Serialize(range, Legacy), wire.Value.GetRawText());
        Assert.Throws<NotSupportedException>(() => WorkerProtocol.EncodePayload(new UnregisteredPayload("hidden")));
    }

    private sealed record UnregisteredPayload(string Value);
}
