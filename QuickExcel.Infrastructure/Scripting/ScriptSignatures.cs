using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Security;

public enum ScriptSignatureStatus { Unsigned, Trusted, UntrustedSelfSigned, Untrusted, Invalid }

public sealed record ScriptSignatureVerification(ScriptSignatureStatus Status, string Message,
    string? Subject = null, string? Issuer = null, string? Fingerprint = null, DateTime? NotAfter = null)
{
    public bool CanAutoEnable => Status == ScriptSignatureStatus.Trusted;
}

/// <summary>Versioned, detached signatures over script semantics, independent of cells and caches.</summary>
public static class ScriptSignatures
{
    public const string CodeSigningOid = "1.3.6.1.5.5.7.3.3";
    private const string RsaAlgorithm = "RSA-PSS-SHA256";
    private const string EcAlgorithm = "ECDSA-P256-SHA256";

    public static ScriptSignature Sign(ScriptModuleSet modules, X509Certificate2 certificate)
    {
        ValidateCertificate(certificate);
        var encoded = Convert.ToBase64String(certificate.RawData);
        byte[] value;
        string algorithm;
        using var rsa = certificate.GetRSAPrivateKey();
        if (rsa is not null)
        {
            if (rsa.KeySize < 2048) throw new CryptographicException("RSA 签名密钥至少需要 2048 位。");
            algorithm = RsaAlgorithm;
            value = rsa.SignData(Payload(modules, algorithm, certificate), HashAlgorithmName.SHA256, RSASignaturePadding.Pss);
        }
        else
        {
            using var ec = certificate.GetECDsaPrivateKey();
            if (ec is null || ec.KeySize != 256 || ec.ExportParameters(false).Curve.Oid.Value != "1.2.840.10045.3.1.7")
                throw new CryptographicException("请选择含 RSA 或 ECDSA P-256 私钥的代码签名证书。");
            algorithm = EcAlgorithm;
            value = ec.SignData(Payload(modules, algorithm, certificate), HashAlgorithmName.SHA256, DSASignatureFormat.IeeeP1363FixedFieldConcatenation);
        }
        if (certificate.RawData.Length > 16 * 1024 || value.Length > 2048) throw new CryptographicException("证书或签名超过容量限制。");
        return new(1, algorithm, encoded, Convert.ToBase64String(value));
    }

    public static ScriptSignatureVerification Verify(ScriptModuleSet modules, IEnumerable<string>? trustedFingerprints = null) =>
        Verify(modules, trustedFingerprints, null);

    internal static ScriptSignatureVerification Verify(ScriptModuleSet modules, IEnumerable<string>? trustedFingerprints,
        Action<X509ChainPolicy>? configureChain)
    {
        if (modules.Signature is not { } signature) return new(ScriptSignatureStatus.Unsigned, "脚本未签名；需要手动运行。");
        try
        {
            if (signature.Version != 1 || signature.Algorithm is not (RsaAlgorithm or EcAlgorithm) ||
                signature.Certificate is null || signature.Certificate.Length > 24 * 1024 || signature.Value is null || signature.Value.Length > 4096)
                throw new CryptographicException("不支持的签名格式或签名超过容量限制。");
            var der = Convert.FromBase64String(signature.Certificate);
            using var certificate = X509CertificateLoader.LoadCertificate(der);
            if (!der.AsSpan().SequenceEqual(certificate.RawData)) throw new CryptographicException("证书必须使用 DER 编码。");
            var value = Convert.FromBase64String(signature.Value);
            var payload = Payload(modules, signature.Algorithm, certificate);
            using var rsa = certificate.GetRSAPublicKey();
            using var ec = certificate.GetECDsaPublicKey();
            var valid = signature.Algorithm == RsaAlgorithm
                ? rsa is { KeySize: >= 2048 } && rsa.VerifyData(payload, value, HashAlgorithmName.SHA256, RSASignaturePadding.Pss)
                : ec is { KeySize: 256 } && ec.ExportParameters(false).Curve.Oid.Value == "1.2.840.10045.3.1.7" &&
                    ec.VerifyData(payload, value, HashAlgorithmName.SHA256, DSASignatureFormat.IeeeP1363FixedFieldConcatenation);
            if (!valid) throw new CryptographicException("脚本签名无效：内容、绑定或证书可能已被更改。");
            ValidateCertificate(certificate);
            var fingerprint = Convert.ToHexString(SHA256.HashData(certificate.RawData));
            ScriptSignatureVerification Result(ScriptSignatureStatus status, string message) => new(status, message,
                DisplayName(certificate.Subject), DisplayName(certificate.Issuer), fingerprint, certificate.NotAfter);
            using var chain = new X509Chain();
            chain.ChainPolicy.ApplicationPolicy.Add(new Oid(CodeSigningOid));
            chain.ChainPolicy.RevocationMode = X509RevocationMode.Online;
            chain.ChainPolicy.RevocationFlag = X509RevocationFlag.ExcludeRoot;
            chain.ChainPolicy.UrlRetrievalTimeout = TimeSpan.FromSeconds(5);
            configureChain?.Invoke(chain.ChainPolicy);
            if (chain.Build(certificate)) return Result(ScriptSignatureStatus.Trusted, "签名有效；证书受系统信任。");
            // Pinning only bypasses an unknown self-signed root. Expiry, revocation, usage,
            // explicit distrust and every other chain failure remain failures even for a pin.
            var selfSigned = certificate.SubjectName.RawData.AsSpan().SequenceEqual(certificate.IssuerName.RawData) &&
                chain.ChainElements.Count == 1 && chain.ChainStatus.Length > 0 &&
                chain.ChainStatus.All(s => s.Status == X509ChainStatusFlags.UntrustedRoot);
            if (selfSigned)
                return trustedFingerprints?.Contains(fingerprint, StringComparer.OrdinalIgnoreCase) == true
                    ? Result(ScriptSignatureStatus.Trusted, "签名有效；已信任此自签证书。")
                    : Result(ScriptSignatureStatus.UntrustedSelfSigned, "签名有效；此自签证书尚未受信任。");
            return Result(ScriptSignatureStatus.Untrusted, "证书链验证失败，未自动启用：" +
                string.Join("、", chain.ChainStatus.Select(s => s.Status.ToString())));
        }
        catch (Exception ex) when (ex is CryptographicException or ArgumentException or FormatException)
        { return new(ScriptSignatureStatus.Invalid, "签名验证失败：" + ex.Message); }
    }

    public static X509Certificate2 CreateSelfSigned(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 128 || name.Any(char.IsControl)) throw new ArgumentException("请输入 1–128 字符的证书名称。");
        using var key = RSA.Create(3072);
        var subject = new X500DistinguishedNameBuilder();
        subject.AddCommonName(name);
        var request = new CertificateRequest(subject.Build(), key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection { new(CodeSigningOid) }, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddYears(2));
    }

    private static void ValidateCertificate(X509Certificate2 certificate)
    {
        var now = DateTime.UtcNow;
        if (now < certificate.NotBefore.ToUniversalTime() || now > certificate.NotAfter.ToUniversalTime())
            throw new CryptographicException("证书已过期或尚未生效；请使用有效证书重新签名。");
        if (!certificate.Extensions.OfType<X509EnhancedKeyUsageExtension>().Any(e => e.EnhancedKeyUsages.Cast<Oid>().Any(o => o.Value == CodeSigningOid)))
            throw new CryptographicException("证书必须声明代码签名用途。");
        if (certificate.Extensions.OfType<X509KeyUsageExtension>().Any(e => !e.KeyUsages.HasFlag(X509KeyUsageFlags.DigitalSignature)))
            throw new CryptographicException("证书不允许数字签名。");
    }

    // BinaryWriter uses length-prefixed strict UTF-8 strings and little-endian Int32 counts.
    // Revisions, cell values and caches are deliberately excluded. Certificate substitution is bound too.
    private static byte[] Payload(ScriptModuleSet modules, string algorithm, X509Certificate2 certificate)
    {
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, new UTF8Encoding(false, true), true))
        {
            writer.Write("QuickExcel.ScriptSignature"); writer.Write(1); writer.Write(algorithm);
            writer.Write(Convert.ToHexString(SHA256.HashData(certificate.RawData)));
            writer.Write(modules.ScriptApiVersion);
            writer.Write(modules.Modules.Count);
            foreach (var (id, module) in modules.Modules.OrderBy(m => m.Key, StringComparer.Ordinal))
            { writer.Write(id); writer.Write(module.Path); writer.Write(module.Source); }
            writer.Write(modules.EntryModules.Length);
            foreach (var entry in modules.EntryModules) writer.Write(entry);
            writer.Write(modules.Bindings.Length);
            foreach (var binding in modules.Bindings.OrderBy(b => b.Module, StringComparer.Ordinal).ThenBy(b => b.Kind, StringComparer.Ordinal))
            { writer.Write(binding.Module); writer.Write("workbook"); writer.Write(binding.Kind); }
        }
        return buffer.ToArray();
    }

    private static string DisplayName(string value) => new(value.Where(c => !char.IsControl(c) &&
        char.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.Format).Take(512).ToArray());
}
