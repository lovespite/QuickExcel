using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using QuickExcel.Domain;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Security;

public sealed record TrustedScriptCertificate(string Fingerprint, string Subject, DateTime NotAfter);

/// <summary>Integrity-protected, per-user QuickExcel trust pins. Never changes Windows certificate stores.</summary>
public sealed class ScriptCertificateTrustStore
{
    private readonly string path;
    private readonly TrustDataProtection _protection;
    private readonly string _executionPolicy;
    public ScriptCertificateTrustStore(string path) : this(path,TrustDataProtection.Create()) { }
    internal ScriptCertificateTrustStore(string path, TrustDataProtection protection, string? executionPolicy = null)
    { this.path = path; _protection = protection; _executionPolicy = executionPolicy ?? CurrentExecutionPolicy; }
    private static string CurrentExecutionPolicy => CompilerProtocol.Hash(JsonSerializer.Serialize(new
    {
        engine = WorkerProtocol.Engine, policy = WorkerProtocol.Policy, protocol = WorkerProtocol.Version,
        compiler = CompilerProtocol.Identity, compilerPolicy = CompilerProtocol.Policy, options = CompilerProtocol.Options
    }));
    private const int MaximumFileSize = 1024 * 1024;
    private sealed record ExecutionConsent(string Fingerprint, string Policy);
    private sealed record TrustState(TrustedScriptCertificate[] Certificates, ExecutionConsent[] ExecutionConsents);
    private sealed record ProtectedTrust(int Version, byte[] Data, [property:JsonIgnore(Condition=JsonIgnoreCondition.WhenWritingNull)] string? Protection=null);
    private static readonly JsonSerializerOptions Format = new() { UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow };

    // Bind the DPAPI payload to this purpose and location, including isolated smoke/test stores.
    // This is domain separation, not a secret. DPAPI is not a boundary against code running as this user.
    private bool IsDpapi=>_protection.Id=="dpapi-current-user-v1";
    private byte[] Entropy => Encoding.UTF8.GetBytes(IsDpapi ? "QuickExcel.ScriptCertificateTrust.v1\0" + Path.GetFullPath(path).ToUpperInvariant()
        : "QuickExcel.ScriptCertificateTrust.v2\0"+_protection.Id+"\0"+Path.GetFullPath(path));

    public IReadOnlyList<TrustedScriptCertificate> Read() => ReadState().Certificates;

    private TrustState ReadState()
    {
        if (!File.Exists(path)) return new([], []);
        using var file = File.OpenRead(path);
        if (file.Length > MaximumFileSize) throw new InvalidDataException("证书信任列表超过容量限制。");
        using var json = JsonDocument.Parse(file, new JsonDocumentOptions { AllowDuplicateProperties = false });
        // Old plaintext pins have no evidence of consent. Never import or protect them on read/write.
        if (json.RootElement.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("旧版或未经验证的证书信任列表不能授予信任；请移除该列表并在应用中重新确认。");
        var envelope = json.RootElement.Deserialize<ProtectedTrust>(Format);
        if (envelope is not { Data.Length: > 0 } || (IsDpapi ? envelope.Version!=1 || envelope.Protection is not null : envelope.Version!=2 || envelope.Protection!=_protection.Id))
            throw new InvalidDataException("证书信任列表格式或平台保护方式不匹配；不会跨平台迁移信任。");
        byte[] plaintext;
        try { plaintext = _protection.Unprotect(envelope.Data, Entropy); }
        catch (CryptographicException ex) { throw new InvalidDataException("证书信任列表完整性验证失败；脚本保持停用。", ex); }
        TrustState state;
        try
        {
            using var payload = JsonDocument.Parse(plaintext, new JsonDocumentOptions { AllowDuplicateProperties = false });
            // Existing protected pins remain valid certificate trust, but grant no engine-policy consent.
            state = payload.RootElement.ValueKind == JsonValueKind.Array
                ? new(payload.RootElement.Deserialize<TrustedScriptCertificate[]>(Format)!, [])
                : payload.RootElement.Deserialize<TrustState>(Format) ?? throw new InvalidDataException("证书信任列表无效。");
        }
        finally { CryptographicOperations.ZeroMemory(plaintext); }
        var entries = state.Certificates;
        var consents = state.ExecutionConsents;
        if (entries is null || entries.Length > 1024 || entries.Any(e => e is null || !ValidHash(e.Fingerprint) || e.Subject is null || e.Subject.Length > 512) ||
            entries.Select(e => e.Fingerprint).Distinct(StringComparer.OrdinalIgnoreCase).Count() != entries.Length ||
            consents is null || consents.Length > 1024 || consents.Any(e => e is null || !ValidHash(e.Fingerprint) || !ValidHash(e.Policy)) ||
            consents.Select(e => e.Fingerprint).Distinct(StringComparer.OrdinalIgnoreCase).Count() != consents.Length)
            throw new InvalidDataException("证书信任列表无效。");
        return state;
    }

    private static bool ValidHash(string? value) => value is { Length: 64 } && value.All(char.IsAsciiHexDigit);

    public ScriptSignatureVerification Verify(ScriptModuleSet modules)
    {
        var state = ReadState();
        var result = ScriptSignatures.Verify(modules, state.Certificates.Select(e => e.Fingerprint));
        return result with { ExecutionPolicyAccepted = result.CanAutoEnable && state.ExecutionConsents.Any(e =>
            e.Fingerprint.Equals(result.Fingerprint, StringComparison.OrdinalIgnoreCase) && e.Policy == _executionPolicy) };
    }

    /// <summary>Called only after explicit Apply; certificate trust alone never accepts a new runtime policy.</summary>
    public void AcceptExecutionPolicy(ScriptModuleSet modules) => Change(state =>
    {
        var result = ScriptSignatures.Verify(modules, state.Certificates.Select(e => e.Fingerprint));
        if (!result.CanAutoEnable) throw new InvalidOperationException("只有签名有效且证书受信任的脚本才能记住执行策略授权。");
        return state with { ExecutionConsents = state.ExecutionConsents.Where(e =>
            !e.Fingerprint.Equals(result.Fingerprint, StringComparison.OrdinalIgnoreCase)).Append(new(result.Fingerprint!, _executionPolicy)).ToArray() };
    });

    public void Trust(ScriptModuleSet modules)
    {
        var result = ScriptSignatures.Verify(modules);
        if (result.Status != ScriptSignatureStatus.UntrustedSelfSigned) throw new InvalidOperationException("只能信任签名有效的自签证书。");
        Change(state => state with { Certificates = state.Certificates.Where(e => !e.Fingerprint.Equals(result.Fingerprint, StringComparison.OrdinalIgnoreCase))
            .Append(new(result.Fingerprint!, result.Subject!, result.NotAfter!.Value)).ToArray() });
    }

    public void Remove(string fingerprint) => Change(state => new(
        state.Certificates.Where(e => !e.Fingerprint.Equals(fingerprint, StringComparison.OrdinalIgnoreCase)).ToArray(),
        state.ExecutionConsents.Where(e => !e.Fingerprint.Equals(fingerprint, StringComparison.OrdinalIgnoreCase)).ToArray()));

    private void Change(Func<TrustState, TrustState> change)
    {
        var full = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        // Serialize read/modify/write across application instances; a failed write never grants trust.
        using var lease = new FileStream(full + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        var state = change(ReadState());
        if (state.Certificates.Length > 1024 || state.ExecutionConsents.Length > 1024) throw new InvalidOperationException("最多信任 1024 张证书，请先移除不用的证书。");
        var plaintext = JsonSerializer.SerializeToUtf8Bytes(state, Format);
        byte[] data;
        try { data = _protection.Protect(plaintext, Entropy); }
        finally { CryptographicOperations.ZeroMemory(plaintext); }
        var encoded = JsonSerializer.SerializeToUtf8Bytes(new ProtectedTrust(IsDpapi?1:2, data,IsDpapi?null:_protection.Id), Format);
        if (encoded.Length > MaximumFileSize) throw new InvalidOperationException("证书信任列表超过容量限制，请先移除不用的证书。");
        var temporary = full + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { file.Write(encoded); file.Flush(true); }
            File.Move(temporary, full, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
