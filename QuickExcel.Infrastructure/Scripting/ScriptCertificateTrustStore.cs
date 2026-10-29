using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using QuickExcel.Domain;

namespace QuickExcel.Security;

public sealed record TrustedScriptCertificate(string Fingerprint, string Subject, DateTime NotAfter);

/// <summary>Integrity-protected, per-user QuickExcel trust pins. Never changes Windows certificate stores.</summary>
public sealed class ScriptCertificateTrustStore(string path)
{
    private const int MaximumFileSize = 1024 * 1024;
    private sealed record ProtectedTrust(int Version, byte[] Data);
    private static readonly JsonSerializerOptions Format = new() { UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow };

    // Bind the DPAPI payload to this purpose and location, including isolated smoke/test stores.
    // This is domain separation, not a secret. DPAPI is not a boundary against code running as this user.
    private byte[] Entropy => Encoding.UTF8.GetBytes("QuickExcel.ScriptCertificateTrust.v1\0" + Path.GetFullPath(path).ToUpperInvariant());

    public IReadOnlyList<TrustedScriptCertificate> Read()
    {
        if (!File.Exists(path)) return [];
        using var file = File.OpenRead(path);
        if (file.Length > MaximumFileSize) throw new InvalidDataException("证书信任列表超过容量限制。");
        using var json = JsonDocument.Parse(file, new JsonDocumentOptions { AllowDuplicateProperties = false });
        // Old plaintext pins have no evidence of consent. Never import or protect them on read/write.
        if (json.RootElement.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("旧版或未经验证的证书信任列表不能授予信任；请移除该列表并在应用中重新确认。");
        var envelope = json.RootElement.Deserialize<ProtectedTrust>(Format);
        if (envelope is not { Version: 1, Data.Length: > 0 }) throw new InvalidDataException("证书信任列表格式无效。");
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("证书信任列表需要 Windows 用户数据保护。");
        byte[] plaintext;
        try { plaintext = ProtectedData.Unprotect(envelope.Data, Entropy, DataProtectionScope.CurrentUser); }
        catch (CryptographicException ex) { throw new InvalidDataException("证书信任列表完整性验证失败；脚本保持停用。", ex); }
        TrustedScriptCertificate[] entries;
        try { entries = JsonSerializer.Deserialize<TrustedScriptCertificate[]>(plaintext, Format) ?? throw new InvalidDataException("证书信任列表无效。"); }
        finally { CryptographicOperations.ZeroMemory(plaintext); }
        if (entries.Length > 1024 || entries.Any(e => e is null || e.Fingerprint is null || e.Fingerprint.Length != 64 ||
            e.Fingerprint.Any(c => !char.IsAsciiHexDigit(c)) || e.Subject is null || e.Subject.Length > 512) ||
            entries.Select(e => e.Fingerprint).Distinct(StringComparer.OrdinalIgnoreCase).Count() != entries.Length)
            throw new InvalidDataException("证书信任列表无效。");
        return entries;
    }

    public ScriptSignatureVerification Verify(ScriptModuleSet modules) => ScriptSignatures.Verify(modules, Read().Select(e => e.Fingerprint));

    public void Trust(ScriptModuleSet modules)
    {
        var result = ScriptSignatures.Verify(modules);
        if (result.Status != ScriptSignatureStatus.UntrustedSelfSigned) throw new InvalidOperationException("只能信任签名有效的自签证书。");
        Change(entries => entries.Where(e => !e.Fingerprint.Equals(result.Fingerprint, StringComparison.OrdinalIgnoreCase))
            .Append(new(result.Fingerprint!, result.Subject!, result.NotAfter!.Value)).ToArray());
    }

    public void Remove(string fingerprint) => Change(entries => entries.Where(e => !e.Fingerprint.Equals(fingerprint, StringComparison.OrdinalIgnoreCase)).ToArray());

    private void Change(Func<IReadOnlyList<TrustedScriptCertificate>, TrustedScriptCertificate[]> change)
    {
        var full = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        // Serialize read/modify/write across application instances; a failed write never grants trust.
        using var lease = new FileStream(full + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        var entries = change(Read());
        if (entries.Length > 1024) throw new InvalidOperationException("最多信任 1024 张证书，请先移除不用的证书。");
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("证书信任列表需要 Windows 用户数据保护。");
        var plaintext = JsonSerializer.SerializeToUtf8Bytes(entries, Format);
        byte[] data;
        try { data = ProtectedData.Protect(plaintext, Entropy, DataProtectionScope.CurrentUser); }
        finally { CryptographicOperations.ZeroMemory(plaintext); }
        var encoded = JsonSerializer.SerializeToUtf8Bytes(new ProtectedTrust(1, data), Format);
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
