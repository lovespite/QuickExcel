using QuickExcel.Serialization;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;
using QuickExcel.Security;

namespace QuickExcel.Runtime;

/// <summary>Disposable derived data keyed by the exact source signature. Source edits drop that identity.
/// Serialized bytes are private, so neither callers nor worker reply arrays can mutate accepted artifacts.</summary>
internal static class ScriptCompilationCache
{
    internal const string Part = "cache/script-compilation.json";
    internal const string DataPart = "cache/script-compilation.bin";
    internal const string Feature = "script-compilation-cache";
    internal const int MaximumMetadataBytes = 64 * 1024;
    internal const int MaximumDataBytes = 5 * 1024 * 1024;
    // Old v1 packages can still be opened; their monolithic JSON is discarded without parsing.
    internal const int LegacyMaximumBytes = 16 * 1024 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private static readonly byte[] Purpose = Encoding.UTF8.GetBytes("QuickExcel.ScriptCompilationCache.v2");
    internal sealed record OutputIndex(string Id, int Offset, int JavascriptBytes, int SourceMapBytes);
    internal sealed record CompilationMetadata(string Compiler, string SourceSha256, string OutputSha256, string SourceMapMode,
        string Frontend, string Mode, double Milliseconds, OutputIndex[] Outputs);
    internal sealed record Envelope(int Version, string SignatureSha256, CompilationMetadata Compilation,
        int DataBytes, string DataSha256, string Protection, string Proof);
    internal sealed record PackageParts(byte[] Metadata, byte[] Data);
    private sealed record Stored(PackageParts Parts, bool ProducedHere);
    private sealed class Slot { internal Stored? Value; }
    private static readonly ConditionalWeakTable<ScriptSignature, Slot> Entries = new();

    internal static void Invalidate(ScriptModuleSet modules)
    {
        if (modules.Signature is { } signature) Entries.Remove(signature);
    }

    internal static void Import(ScriptModuleSet modules, byte[] metadata, byte[] data)
    {
        if (modules.Signature is { } signature && metadata.Length <= MaximumMetadataBytes && data.Length <= MaximumDataBytes)
            Volatile.Write(ref Entries.GetOrCreateValue(signature).Value, new(new(metadata.ToArray(), data.ToArray()), false));
    }

    internal static PackageParts? Export(ScriptModuleSet modules)
    {
        if (!TryRead(modules, ScriptCompilation.Request(modules), out _)) return null;
        var stored = Volatile.Read(ref Entries.GetOrCreateValue(modules.Signature!).Value);
        return stored is null ? null : new(stored.Parts.Metadata.ToArray(), stored.Parts.Data.ToArray());
    }

    internal static bool TryRead(ScriptModuleSet modules, CompileRequest request, out CompiledReply? result)
    {
        result = null;
        if (modules.Signature is not { } signature || !Entries.TryGetValue(signature, out var slot) || Volatile.Read(ref slot.Value) is not { } stored)
            return false;
        try
        {
            // Only the small metadata document is parsed. Check all offsets/lengths and authenticate
            // the binary digest before allocating any executable strings from untrusted data.
            using var json = JsonDocument.Parse(stored.Parts.Metadata, new() { AllowDuplicateProperties = false, MaxDepth = 16 });
            var cache = json.RootElement.Deserialize(InfrastructureJsonContext.Protocol.CompilationEnvelope);
            if (cache is not { Version: 2, Compilation: not null } || cache.SignatureSha256 != SignatureHash(signature) ||
                !ValidIndex(cache, request, stored.Parts.Data.Length) || !ScriptSignatures.VerifyContent(modules) ||
                cache.DataSha256 != Convert.ToHexString(SHA256.HashData(stored.Parts.Data))) return false;
            var payload = Payload(cache);
            if (!stored.ProducedHere)
            {
                if (cache.Proof is null || cache.Proof.Length > 16384) return false;
                var proof = Convert.FromBase64String(cache.Proof);
                if (cache.Protection == "publisher")
                {
                    using var certificate = X509CertificateLoader.LoadCertificate(Convert.FromBase64String(signature.Certificate));
                    if (!ScriptSignatures.VerifyData(payload, proof, certificate, signature.Algorithm)) return false;
                }
                else
                {
                    var protection = TrustDataProtection.Create();
                    if (cache.Protection != protection.Id) return false;
                    var verified = protection.Unprotect(proof, Purpose);
                    try { if (!CryptographicOperations.FixedTimeEquals(payload, verified)) return false; }
                    finally { CryptographicOperations.ZeroMemory(verified); }
                }
            }
            var outputs = cache.Compilation.Outputs.Select(o => new CompilerOutput(o.Id,
                Utf8.GetString(stored.Parts.Data.AsSpan(o.Offset, o.JavascriptBytes)),
                Utf8.GetString(stored.Parts.Data.AsSpan(o.Offset + o.JavascriptBytes, o.SourceMapBytes)))).ToArray();
            var compiled = new CompiledReply(cache.Compilation.Compiler, cache.Compilation.SourceSha256, cache.Compilation.OutputSha256,
                new(cache.Compilation.Frontend, cache.Compilation.Mode, true, cache.Compilation.Milliseconds, [], outputs));
            CompilerProtocol.ValidateReply(compiled, request);
            result = compiled;
            return true;
        }
        catch (Exception ex) when (ex is JsonException or IOException or CryptographicException or ArgumentException or
            FormatException or InvalidOperationException or KeyNotFoundException or OverflowException or NotSupportedException)
        { return false; } // A disposable cache never prevents recovery from the original signed sources.
    }

    internal static void StoreSigned(ScriptModuleSet modules, CompileRequest request, CompiledReply compiled, X509Certificate2 certificate, CancellationToken token)
    {
        var (envelope, data) = Create(modules, request, compiled, "publisher");
        Store(modules, envelope with { Proof = Convert.ToBase64String(ScriptSignatures.SignData(Payload(envelope), certificate, modules.Signature!.Algorithm)) }, data, token);
    }

    internal static void StoreLocal(ScriptModuleSet modules, CompileRequest request, CompiledReply compiled, CancellationToken token)
    {
        var (envelope, data) = Create(modules, request, compiled, "session");
        try
        {
            var protection = TrustDataProtection.Create();
            envelope = envelope with { Protection = protection.Id };
            envelope = envelope with { Proof = Convert.ToBase64String(protection.Protect(Payload(envelope), Purpose)) };
        }
        catch (Exception ex) when (ex is IOException or CryptographicException or PlatformNotSupportedException)
        { envelope = envelope with { Protection = "session", Proof = "" }; }
        // No publisher private key is needed to rebuild. On another OS account/device the local
        // receipt fails closed and the original signed sources are compiled again.
        Store(modules, envelope, data, token);
    }

    private static bool ValidIndex(Envelope cache, CompileRequest request, int dataLength)
    {
        var metadata = cache.Compilation;
        if (cache.DataBytes != dataLength || dataLength > MaximumDataBytes || metadata.Compiler != CompilerProtocol.Identity ||
            metadata.SourceSha256 != request.Sha256 || metadata.SourceMapMode != request.SourceMapMode || metadata.Frontend != "5.9.3" || metadata.Mode != "check" ||
            !double.IsFinite(metadata.Milliseconds) || metadata.Milliseconds < 0 || metadata.Outputs is null ||
            metadata.Outputs.Length is < 1 or > 64 || metadata.Outputs.Length != request.Modules.Length) return false;
        var ids = new HashSet<string>(StringComparer.Ordinal);
        long offset = 0, javascriptBytes = 0, mapBytes = 0;
        foreach (var output in metadata.Outputs)
        {
            if (output is null || !CompilerProtocol.ValidId(output.Id) || !ids.Add(output.Id) ||
                !request.Modules.Any(m => m.Id == output.Id) || output.Offset != offset || output.JavascriptBytes < 0 || output.SourceMapBytes < 0)
                return false;
            javascriptBytes += output.JavascriptBytes; mapBytes += output.SourceMapBytes;
            offset += (long)output.JavascriptBytes + output.SourceMapBytes;
            if (javascriptBytes > 1024 * 1024 || mapBytes > 4 * 1024 * 1024 || offset > dataLength) return false;
        }
        return offset == dataLength;
    }

    private static (Envelope Metadata, byte[] Data) Create(ScriptModuleSet modules, CompileRequest request, CompiledReply compiled, string protection)
    {
        CompilerProtocol.ValidateReply(compiled, request);
        if (!compiled.Result.Success || modules.Signature is null) throw new InvalidOperationException("Cannot cache unsigned or failed compilation.");
        var dataLength = compiled.Result.Outputs.Sum(o => Utf8.GetByteCount(o.Javascript) + Utf8.GetByteCount(o.SourceMap));
        var data = new byte[dataLength];
        var offset = 0;
        var outputs = compiled.Result.Outputs.Select(output =>
        {
            var start = offset;
            var javascriptBytes = Utf8.GetBytes(output.Javascript, data.AsSpan(offset)); offset += javascriptBytes;
            var sourceMapBytes = Utf8.GetBytes(output.SourceMap, data.AsSpan(offset)); offset += sourceMapBytes;
            return new OutputIndex(output.Id, start, javascriptBytes, sourceMapBytes);
        }).ToArray();
        return (new(2, SignatureHash(modules.Signature), new(compiled.Compiler, compiled.SourceSha256, compiled.OutputSha256, request.SourceMapMode,
            compiled.Result.Frontend, compiled.Result.Mode, compiled.Result.Milliseconds, outputs),
            data.Length, Convert.ToHexString(SHA256.HashData(data)), protection, ""), data);
    }

    private static void Store(ScriptModuleSet modules, Envelope envelope, byte[] data, CancellationToken token)
    {
        var metadata = JsonSerializer.SerializeToUtf8Bytes(envelope, InfrastructureJsonContext.Protocol.CompilationEnvelope);
        if (metadata.Length > MaximumMetadataBytes || data.Length > MaximumDataBytes) throw new InvalidOperationException(QuickExcel.Core.MessageText.Get("脚本编译缓存超过容量限额。"));
        token.ThrowIfCancellationRequested();
        Volatile.Write(ref Entries.GetOrCreateValue(modules.Signature!).Value, new(new(metadata, data), true));
    }

    private static string SignatureHash(ScriptSignature signature) => CompilerProtocol.Hash(JsonSerializer.Serialize(signature, InfrastructureJsonContext.Protocol.ScriptSignature));
    // Authenticate metadata and the binary digest. JS and source maps never enter this JSON payload.
    internal sealed record Authentication(string Purpose, int Version, string SignatureSha256, CompilationMetadata Compilation,
        int DataBytes, string DataSha256, string Protection);
    private static byte[] Payload(Envelope envelope) => SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(new Authentication(
        "QuickExcel.ScriptCompilationCache", envelope.Version, envelope.SignatureSha256, envelope.Compilation,
        envelope.DataBytes, envelope.DataSha256, envelope.Protection), InfrastructureJsonContext.Protocol.Authentication));
}
