using System.Text.Json;
using QuickExcel.Domain;

namespace QuickExcel.Security;

public sealed record TrustedScriptCertificate(string Fingerprint, string Subject, DateTime NotAfter);

/// <summary>Per-user QuickExcel trust pins. Never adds certificates to the Windows root store.</summary>
public sealed class ScriptCertificateTrustStore(string path)
{
    public IReadOnlyList<TrustedScriptCertificate> Read()
    {
        if (!File.Exists(path)) return [];
        using var file = File.OpenRead(path);
        if (file.Length > 1024 * 1024) throw new InvalidDataException("证书信任列表超过容量限制。");
        var entries = JsonSerializer.Deserialize<TrustedScriptCertificate[]>(file) ?? throw new InvalidDataException("证书信任列表无效。");
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
        var temporary = full + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { JsonSerializer.Serialize(file, entries); file.Flush(true); }
            File.Move(temporary, full, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
