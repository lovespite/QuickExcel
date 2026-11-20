using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace QuickExcel.ScriptContracts;

public sealed record CompilerModule(string Id, string Language, string Source);
public sealed record CompileRequest(CompilerModule[] Modules, string[] Entries, string Sha256)
{
    public int ApiVersion { get; init; } = 3;
    public string SourceMapMode { get; init; } = "full";
    public string ComputeDigest() => CompilerProtocol.Hash(JsonSerializer.Serialize(new CompilationDigest(
        CompilerProtocol.Identity, CompilerProtocol.Policy, CompilerProtocol.Options,
        WorkerProtocol.Engine, WorkerProtocol.Policy, ApiVersion, SourceMapMode,
        Modules.OrderBy(m => m.Id, StringComparer.Ordinal).ToArray(), Entries), DigestJsonContext.Default.CompilationDigest));

    public void Validate()
    {
        if (ApiVersion is not (3 or 4) || SourceMapMode is not ("none" or "compact" or "full") || Modules.Length is < 1 or > 64 || Entries.Length is < 1 or > 64 ||
            Modules.Any(m => m is null || !CompilerProtocol.ValidId(m.Id) || m.Language is not ("javascript" or "typescript") || m.Source is null) ||
            Modules.Select(m => m.Id).Distinct(StringComparer.Ordinal).Count() != Modules.Length ||
            Entries.Distinct(StringComparer.Ordinal).Count() != Entries.Length || Entries.Any(id => !Modules.Any(m => m.Id == id)) ||
            Modules.Sum(m => (long)Encoding.UTF8.GetByteCount(m.Source)) > 1024 * 1024 || Sha256 != ComputeDigest())
            throw new ProtocolException("Invalid compilation source identity or budget.");
    }
}

public sealed record CompilerDiagnostic(JsonElement Code, string Message, string? Module, int? Line, int? Column);
public sealed record CompilerOutput(string Id, string Javascript, string SourceMap);
public sealed record CompilerResult(string Frontend, string Mode, bool Success, double Milliseconds,
    CompilerDiagnostic[] Diagnostics, CompilerOutput[] Outputs);
public sealed record CompiledReply(string Compiler, string SourceSha256, string OutputSha256, CompilerResult Result);

public static class CompilerProtocol
{
    public const string Options = "ts5.9.3-es2022-esnext-strict-isolated-verbatim-define-v1";
    public const string Policy = "typescript-check-1024m-5s-v1";
    public static string Identity { get; } = ReadIdentity();
    public static WorkerHello Hello => new(WorkerProtocol.Engine, Policy, [3, 4], WorkerProtocol.MaxFrameBytes, "compile", Identity);
    private static string ReadIdentity()
    {
        using var stream = typeof(CompilerProtocol).Assembly.GetManifestResourceStream("QuickExcel.Compiler.Identity")
            ?? throw new InvalidOperationException("Compiler build identity is missing.");
        using var reader = new StreamReader(stream);
        var value = reader.ReadToEnd().Trim();
        if (value.Length != 64 || value.Any(c => !char.IsAsciiHexDigitUpper(c))) throw new InvalidOperationException("Invalid compiler build identity.");
        return value;
    }
    public static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    public static bool ValidId(string id) => !string.IsNullOrEmpty(id) && id.Length <= 128 && id is not ("." or "..") &&
        id.All(c => char.IsAsciiLetterOrDigit(c) || c is '_' or '-' or '.');
    public static string OutputDigest(CompilerResult result) => Hash(JsonSerializer.Serialize(result.Outputs, WorkerProtocol.TypeInfo<CompilerOutput[]>()));

    public static void ValidateReply(CompiledReply reply, CompileRequest request)
    {
        var result = reply.Result;
        if (reply.Compiler != Identity || reply.SourceSha256 != request.Sha256 || result is null ||
            result.Frontend != "5.9.3" || result.Mode != "check" || !double.IsFinite(result.Milliseconds) || result.Milliseconds < 0 ||
            result.Diagnostics is null || result.Outputs is null || result.Diagnostics.Length > 128 || result.Outputs.Length > 64 ||
            result.Success != (result.Diagnostics.Length == 0) || !result.Success && result.Outputs.Length != 0 ||
            result.Success && result.Outputs.Length != request.Modules.Length ||
            Encoding.UTF8.GetByteCount(JsonSerializer.Serialize(result.Diagnostics, WorkerProtocol.TypeInfo<CompilerDiagnostic[]>())) > 16 * 1024 ||
            result.Diagnostics.Any(d => d is null || d.Message is null || d.Message.Length > 4200 ||
                d.Module is not null && !request.Modules.Any(m => m.Id == d.Module) ||
                d.Line is <= 0 || d.Column is <= 0 || d.Line.HasValue != d.Column.HasValue ||
                d.Code.ValueKind is not (JsonValueKind.String or JsonValueKind.Number)) ||
            result.Outputs.Any(o => o is null || o.Javascript is null || o.SourceMap is null) ||
            result.Outputs.Select(o => o.Id).Distinct(StringComparer.Ordinal).Count() != result.Outputs.Length ||
            result.Outputs.Sum(o => (long)Encoding.UTF8.GetByteCount(o.Javascript)) > 1024 * 1024 ||
            result.Outputs.Sum(o => (long)Encoding.UTF8.GetByteCount(o.SourceMap)) > 4 * 1024 * 1024 ||
            reply.OutputSha256 != OutputDigest(result))
            throw new ProtocolException("Compiler reply identity, shape or byte budget is invalid.");
        foreach (var output in result.Outputs)
        {
            var input = request.Modules.SingleOrDefault(m => m.Id == output.Id)
                ?? throw new ProtocolException("Unknown compiler output module.");
            if (request.SourceMapMode == "none")
            {
                if (output.SourceMap.Length != 0) throw new ProtocolException("Source maps are disabled for this compilation.");
                continue;
            }
            using var map = JsonDocument.Parse(output.SourceMap, new() { MaxDepth = 8 });
            var root = map.RootElement;
            var keys = new HashSet<string>(StringComparer.Ordinal);
            if (root.ValueKind != JsonValueKind.Object || root.EnumerateObject().Any(p => !keys.Add(p.Name) || p.Name is not
                ("version" or "file" or "sourceRoot" or "sources" or "sourcesContent" or "names" or "mappings")) ||
                keys.Count != (request.SourceMapMode == "full" ? 7 : 6) || root.GetProperty("version").GetInt32() != 3 ||
                root.GetProperty("file").GetString() != output.Id + ".js" || root.GetProperty("sourceRoot").GetString() != "" ||
                root.GetProperty("sources").GetArrayLength() != 1 || root.GetProperty("sources")[0].GetString() != output.Id ||
                (request.SourceMapMode == "full"
                    ? root.GetProperty("sourcesContent").GetArrayLength() != 1 || root.GetProperty("sourcesContent")[0].GetString() != input.Source
                    : root.TryGetProperty("sourcesContent", out _)) ||
                root.GetProperty("names").ValueKind != JsonValueKind.Array ||
                root.GetProperty("names").EnumerateArray().Any(n => n.ValueKind != JsonValueKind.String) ||
                root.GetProperty("mappings").ValueKind != JsonValueKind.String)
                throw new ProtocolException("Compiler source map escaped its logical source or has an invalid shape.");
            _ = new LogicalSourceMap(root.GetProperty("mappings").GetString()!, input.Source, root.GetProperty("names").GetArrayLength());
        }
    }
}
