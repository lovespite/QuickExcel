using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Domain;
using QuickExcel.Security;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ScriptSignatureTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-signatures-" + Guid.NewGuid().ToString("N"));
    private readonly X509Certificate2 _certificate = ScriptSignatures.CreateSelfSigned("签名测试");
    private static ScriptModuleSet Modules(int api = 2) => new(new Dictionary<string, ScriptModule>
    {
        ["main"] = new("scripts/main.js", "import {value} from 'helper';\nconsole.log(value);"),
        ["helper"] = new("scripts/helper.js", "export const value = '中文';")
    }, ["main", "helper"], scriptApiVersion: api);
    private ScriptModuleSet Signed(ScriptModuleSet? modules = null)
    { modules ??= Modules(); return modules.WithSignature(ScriptSignatures.Sign(modules, _certificate)); }
    private static WorkbookSnapshot Book(ScriptModuleSet modules) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new WorksheetSnapshot(SheetId.New(), "Sheet1")], modules.PrimarySource, scriptModules: modules);
    private string PathFor(string name) => Path.Combine(_folder, name);
    public ScriptSignatureTests() => Directory.CreateDirectory(_folder);

    [Fact]
    public void SelfSignedNeedsExplicitTrustAndTrustCanBeRemoved()
    {
        var modules = Signed(); var store = new ScriptCertificateTrustStore(PathFor("trust.json"));
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, store.Verify(modules).Status);
        store.Trust(modules);
        var reopened = new ScriptCertificateTrustStore(PathFor("trust.json"));
        Assert.True(reopened.Verify(modules).CanAutoEnable);
        var other = Modules().WithPrimarySource("different source");
        other = Signed(other);
        Assert.True(reopened.Verify(other).CanAutoEnable);
        using var otherCert = ScriptSignatures.CreateSelfSigned("签名测试");
        Assert.False(reopened.Verify(other.WithSignature(ScriptSignatures.Sign(other, otherCert))).CanAutoEnable);
        reopened.Remove(reopened.Read().Single().Fingerprint);
        Assert.False(store.Verify(modules).CanAutoEnable);
    }

    [Fact]
    public void ForgedPlaintextTrustListCannotAuthorizeScripts()
    {
        var modules = Signed();
        var path = PathFor("trust.json");
        var forged = System.Text.Json.JsonSerializer.Serialize(new[]
        {
            new TrustedScriptCertificate(Convert.ToHexString(SHA256.HashData(_certificate.RawData)),
                _certificate.Subject, _certificate.NotAfter)
        });
        File.WriteAllText(path, forged);
        var store = new ScriptCertificateTrustStore(path);
        Assert.Throws<InvalidDataException>(() => store.Verify(modules));
        Assert.Throws<InvalidDataException>(() => store.Trust(modules));
        Assert.Equal(forged, File.ReadAllText(path));
    }

    [Theory]
    [InlineData("replace-with-plaintext")]
    [InlineData("modify-protected-data")]
    [InlineData("append-unprotected-pin")]
    [InlineData("copy-from-other-store")]
    public void TamperingWithPersistedTrustCannotAuthorizeAnotherSigner(string mutation)
    {
        var path = PathFor("trust.json");
        var store = new ScriptCertificateTrustStore(path);
        store.Trust(Signed());
        using var attacker = ScriptSignatures.CreateSelfSigned("Injected signer");
        var modules = Modules();
        modules = modules.WithSignature(ScriptSignatures.Sign(modules, attacker));
        var pin = new TrustedScriptCertificate(Convert.ToHexString(SHA256.HashData(attacker.RawData)), attacker.Subject, attacker.NotAfter);
        var original = File.ReadAllText(path);
        var envelope = JsonNode.Parse(original)!.AsObject();
        switch (mutation)
        {
            case "replace-with-plaintext":
                File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new[] { pin }));
                break;
            case "modify-protected-data":
                var data = Convert.FromBase64String(envelope["Data"]!.GetValue<string>());
                data[^1] ^= 1;
                envelope["Data"] = Convert.ToBase64String(data);
                File.WriteAllText(path, envelope.ToJsonString());
                break;
            case "append-unprotected-pin":
                envelope["Entries"] = System.Text.Json.JsonSerializer.SerializeToNode(new[] { pin });
                File.WriteAllText(path, envelope.ToJsonString());
                break;
            case "copy-from-other-store":
                var otherPath = PathFor("attacker.json");
                new ScriptCertificateTrustStore(otherPath).Trust(modules);
                File.Copy(otherPath, path, true);
                break;
        }
        var tampered = File.ReadAllText(path);
        var reopened = new ScriptCertificateTrustStore(path);
        // Verify and an explicit trust write must both fail, without laundering existing forged entries.
        if (mutation == "append-unprotected-pin")
        {
            Assert.ThrowsAny<System.Text.Json.JsonException>(() => reopened.Verify(modules));
            Assert.ThrowsAny<System.Text.Json.JsonException>(() => reopened.Trust(Signed()));
        }
        else
        {
            Assert.Throws<InvalidDataException>(() => reopened.Verify(modules));
            Assert.Throws<InvalidDataException>(() => reopened.Trust(Signed()));
        }
        Assert.Equal(tampered, File.ReadAllText(path));
        File.WriteAllText(path, original);
        Assert.True(reopened.Verify(Signed()).CanAutoEnable);
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, reopened.Verify(modules).Status);
    }

    [Fact]
    public void DeletedTrustFileRequiresConfirmationAgain()
    {
        var modules = Signed(); var path = PathFor("trust.json");
        var store = new ScriptCertificateTrustStore(path); store.Trust(modules);
        File.Delete(path);
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, new ScriptCertificateTrustStore(path).Verify(modules).Status);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RevokedCertificateCannotAutoEnableEvenWithMatchingPin(bool pinned)
    {
        // A fresh CA and loopback CRL exercise Windows' real Online revocation check without
        // installing a root, contacting a public CA, or simulating the chain result.
        using var server = new CrlServer();
        var now = DateTimeOffset.UtcNow;
        using var issuerKey = RSA.Create(2048);
        var issuerRequest = new CertificateRequest("CN=Revocation test " + Guid.NewGuid().ToString("N"), issuerKey,
            HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        issuerRequest.CertificateExtensions.Add(new X509BasicConstraintsExtension(true, false, 0, true));
        issuerRequest.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.KeyCertSign | X509KeyUsageFlags.CrlSign, true));
        issuerRequest.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(issuerRequest.PublicKey, false));
        using var issuer = issuerRequest.CreateSelfSigned(now.AddDays(-1), now.AddDays(5));
        using var key = RSA.Create(2048);
        var request = new CertificateRequest("CN=Revoked code signer", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection { new(ScriptSignatures.CodeSigningOid) }, true));
        request.CertificateExtensions.Add(X509AuthorityKeyIdentifierExtension.CreateFromCertificate(issuer, true, false));
        request.CertificateExtensions.Add(CertificateRevocationListBuilder.BuildCrlDistributionPointExtension([server.Url]));
        using var publicCertificate = request.Create(issuer, now.AddHours(-1), now.AddDays(1), RandomNumberGenerator.GetBytes(16));
        using var certificate = publicCertificate.CopyWithPrivateKey(key);
        using var goodPublic = request.Create(issuer, now.AddHours(-1), now.AddDays(1), RandomNumberGenerator.GetBytes(16));
        using var goodCertificate = goodPublic.CopyWithPrivateKey(key);
        var crl = new CertificateRevocationListBuilder();
        crl.AddEntry(certificate, now.AddMinutes(-10), X509RevocationReason.KeyCompromise);
        server.Data = crl.Build(issuer, 1, now.AddDays(1), HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1, now.AddMinutes(-5));
        void Configure(X509ChainPolicy policy)
        {
            policy.TrustMode = X509ChainTrustMode.CustomRootTrust;
            policy.CustomTrustStore.Add(issuer);
        }
        var unsigned = Modules();
        var good = unsigned.WithSignature(ScriptSignatures.Sign(unsigned, goodCertificate));
        var goodResult = ScriptSignatures.Verify(good, null, Configure);
        Assert.True(goodResult.CanAutoEnable, goodResult.Message);
        var signed = unsigned.WithSignature(ScriptSignatures.Sign(unsigned, certificate));
        string[]? pins = pinned ? [Convert.ToHexString(SHA256.HashData(certificate.RawData))] : null;
        var result = ScriptSignatures.Verify(signed, pins, Configure);
        Assert.Equal(ScriptSignatureStatus.Untrusted, result.Status);
        Assert.False(result.CanAutoEnable);
        Assert.Contains("Revoked", result.Message);
        Assert.Throws<InvalidOperationException>(() => new ScriptCertificateTrustStore(PathFor("trust.json")).Trust(signed));
    }

    private sealed class CrlServer : IDisposable
    {
        private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
        private readonly CancellationTokenSource _stop = new();
        private readonly Task _serve;
        public string Url { get; }
        public byte[] Data { get; set; } = [];

        public CrlServer()
        {
            _listener.Start();
            Url = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/{Guid.NewGuid():N}.crl";
            _serve = ServeAsync();
        }

        private async Task ServeAsync()
        {
            try
            {
                while (!_stop.IsCancellationRequested)
                {
                    using var client = await _listener.AcceptTcpClientAsync(_stop.Token);
                    await using var stream = client.GetStream();
                    using var reader = new StreamReader(stream, Encoding.ASCII, leaveOpen: true);
                    while (await reader.ReadLineAsync(_stop.Token) is { Length: > 0 }) { }
                    var data = Data;
                    var header = Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Type: application/pkix-crl\r\nContent-Length: {data.Length}\r\nConnection: close\r\n\r\n");
                    await stream.WriteAsync(header, _stop.Token);
                    await stream.WriteAsync(data, _stop.Token);
                }
            }
            catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
            catch (SocketException) when (_stop.IsCancellationRequested) { }
        }

        public void Dispose()
        {
            _stop.Cancel(); _listener.Stop();
            try { _serve.GetAwaiter().GetResult(); }
            finally { _stop.Dispose(); }
        }
    }

    [Theory]
    [InlineData("source")]
    [InlineData("helper")]
    [InlineData("path")]
    [InlineData("id")]
    [InlineData("api")]
    [InlineData("entries")]
    [InlineData("bindings")]
    [InlineData("added")]
    [InlineData("description")]
    public void EveryExecutablePartIsCovered(string mutation)
    {
        var signed = Signed(); var map = signed.Modules; var entries = signed.EntryModules.ToArray();
        var bindings = signed.Bindings.ToArray(); var api = signed.ScriptApiVersion;
        switch (mutation)
        {
            case "source": map = map.SetItem("main", map["main"] with { Source = "console.log('evil');" }); break;
            case "helper": map = map.SetItem("helper", map["helper"] with { Source = "export const value = 'other';" }); break;
            case "path": map = map.SetItem("helper", map["helper"] with { Path = "scripts/other.js" }); break;
            case "description": map = map.SetItem("main", map["main"] with { Description = "误导性描述" }); break;
            case "id": map = map.Remove("helper").Add("other", map["helper"]); entries = ["main", "other"]; bindings = [new("main", "formula-functions"), new("other", "formula-functions")]; break;
            case "api": api = 1; break;
            case "entries": entries = ["helper", "main"]; break;
            case "bindings": bindings = [new("main", "actions"), new("helper", "formula-functions")]; break;
            case "added": map = map.Add("unused", new("scripts/unused.js", "")); break;
        }
        var changed = new ScriptModuleSet(map, entries, scriptApiVersion: api, bindings: bindings, signature: signed.Signature);
        Assert.Equal(ScriptSignatureStatus.Invalid, ScriptSignatures.Verify(changed).Status);
    }

    [Fact]
    public void CanonicalPayloadIgnoresMapOrderAndRevisionButBindsCertificate()
    {
        var signed = Signed();
        var reordered = new ScriptModuleSet(signed.Modules.Reverse(), signed.EntryModules, 99, 8, 2, signed.Bindings.Reverse(), signed.Signature);
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, ScriptSignatures.Verify(reordered).Status);
        using var replacement = ScriptSignatures.CreateSelfSigned("replacement");
        Assert.Equal(ScriptSignatureStatus.Invalid, ScriptSignatures.Verify(signed.WithSignature(signed.Signature! with { Certificate = Convert.ToBase64String(replacement.RawData) })).Status);
    }

    [Fact]
    public void ValidSystemChainPolicyAutomaticallyTrustsCodeSigningCertificate()
    {
        // An isolated custom root exercises the same chain branch without changing Windows stores.
        var result = ScriptSignatures.Verify(Signed(), null, p =>
        { p.TrustMode = X509ChainTrustMode.CustomRootTrust; p.CustomTrustStore.Add(_certificate); p.RevocationMode = X509RevocationMode.NoCheck; });
        Assert.True(result.CanAutoEnable);
    }

    [Fact]
    public void EcdsaAndLegacyApiCanBeSigned()
    {
        using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var request = new CertificateRequest("CN=EC", key, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection { new(ScriptSignatures.CodeSigningOid) }, false));
        using var cert = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1));
        var modules = ScriptModuleSet.Single("host.functions.register('F', () => 1);");
        var signed = modules.WithSignature(ScriptSignatures.Sign(modules, cert));
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, ScriptSignatures.Verify(signed).Status);
    }

    [Fact]
    public void ATrustPinCannotOverrideCertificateExpiryOrOtherChainFailures()
    {
        var signed = Signed();
        var fingerprint = ScriptSignatures.Verify(signed).Fingerprint!;
        var expired = ScriptSignatures.Verify(signed, [fingerprint], p =>
        { p.VerificationTime = _certificate.NotAfter.AddDays(1); p.RevocationMode = X509RevocationMode.NoCheck; });
        Assert.False(expired.CanAutoEnable);
        Assert.NotEqual(ScriptSignatureStatus.UntrustedSelfSigned, expired.Status);

        using var key = RSA.Create(2048);
        var request = new CertificateRequest("CN=Unknown critical extension", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection { new(ScriptSignatures.CodeSigningOid) }, false));
        request.CertificateExtensions.Add(new X509Extension("1.2.3.4.5.6.7.8", [5, 0], true));
        using var cert = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1));
        var modules = Modules(); modules = modules.WithSignature(ScriptSignatures.Sign(modules, cert));
        var result = ScriptSignatures.Verify(modules, [Convert.ToHexString(SHA256.HashData(cert.RawData))]);
        Assert.False(result.CanAutoEnable);
        Assert.NotEqual(ScriptSignatureStatus.UntrustedSelfSigned, result.Status);
    }

    [Fact]
    public void UnknownIssuerCannotBeTreatedAsSelfSignedOrPinned()
    {
        using var issuerKey = RSA.Create(2048);
        var issuerRequest = new CertificateRequest("CN=Unknown root", issuerKey, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        issuerRequest.CertificateExtensions.Add(new X509BasicConstraintsExtension(true, false, 0, true));
        using var issuer = issuerRequest.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(5));
        using var key = RSA.Create(2048);
        var request = new CertificateRequest("CN=Leaf", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection { new(ScriptSignatures.CodeSigningOid) }, false));
        using var publicCert = request.Create(issuer, DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1), RandomNumberGenerator.GetBytes(16));
        using var cert = publicCert.CopyWithPrivateKey(key);
        var modules = Modules(); modules = modules.WithSignature(ScriptSignatures.Sign(modules, cert));
        Assert.Equal(ScriptSignatureStatus.Untrusted, ScriptSignatures.Verify(modules, [Convert.ToHexString(SHA256.HashData(cert.RawData))]).Status);
        Assert.Throws<InvalidOperationException>(() => new ScriptCertificateTrustStore(PathFor("trust.json")).Trust(modules));
    }

    [Theory]
    [InlineData("expired")]
    [InlineData("future")]
    [InlineData("server-auth")]
    [InlineData("key-usage")]
    [InlineData("weak-key")]
    public void SigningRejectsUnsuitableCertificates(string kind)
    {
        using var key = RSA.Create(kind == "weak-key" ? 1024 : 2048);
        var request = new CertificateRequest("CN=bad", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(new OidCollection { new(kind == "server-auth" ? "1.3.6.1.5.5.7.3.1" : ScriptSignatures.CodeSigningOid) }, false));
        if (kind == "key-usage") request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.KeyEncipherment, true));
        var start = kind == "future" ? DateTimeOffset.UtcNow.AddDays(1) : DateTimeOffset.UtcNow.AddDays(-2);
        var end = kind == "expired" ? DateTimeOffset.UtcNow.AddDays(-1) : DateTimeOffset.UtcNow.AddDays(2);
        using var cert = request.CreateSelfSigned(start, end);
        Assert.Throws<CryptographicException>(() => ScriptSignatures.Sign(Modules(), cert));
    }

    [Fact]
    public void InvalidMetadataCannotBeTrustedEvenWithPinnedCertificate()
    {
        var signed = Signed(); var store = new ScriptCertificateTrustStore(PathFor("trust.json")); store.Trust(signed);
        foreach (var signature in new[] { signed.Signature! with { Version = 9 }, signed.Signature! with { Algorithm = "none" },
            signed.Signature! with { Certificate = "not-base64" }, signed.Signature! with { Value = "" }, signed.Signature! with { Value = new('A', 5000) } })
        {
            var invalid = signed.WithSignature(signature);
            Assert.Equal(ScriptSignatureStatus.Invalid, store.Verify(invalid).Status);
            Assert.Throws<InvalidOperationException>(() => store.Trust(invalid));
        }
        Assert.False(store.Verify(Modules()).CanAutoEnable);
    }

    [Fact]
    public void SessionKeepsSignatureAcrossCellsAndHistoryButDropsItOnScriptChanges()
    {
        using var session = new DocumentSession(Book(Signed()));
        var signature = session.Snapshot.ScriptModules.Signature;
        var address = new CellAddress(session.Snapshot.Sheets[0].Id, 0, 0);
        session.Execute(new SetCellsCommand([new(address, CellContent.FromLiteral(CellValue.FromNumber(3)))]));
        session.Undo(); session.Redo();
        Assert.Equal(signature, session.Snapshot.ScriptModules.Signature);
        session.SetScriptsEnabled(true);
        session.Execute(new UpdateScriptSourceCommand("new source"));
        Assert.False(session.ScriptsEnabled); Assert.Null(session.Snapshot.ScriptModules.Signature);
        session.Undo(); Assert.Null(session.Snapshot.ScriptModules.Signature);
        session.Redo(); Assert.Null(session.Snapshot.ScriptModules.Signature);
        Assert.Equal("new source", session.Snapshot.ScriptSource);
    }

    [Fact]
    public void SignatureAttachmentIsDirtyNonUndoableAndRejectsStaleSourceCapture()
    {
        using var session = new DocumentSession(Book(Modules()));
        var modules = session.Snapshot.ScriptModules; var signature = ScriptSignatures.Sign(modules, _certificate);
        session.SetScriptsEnabled(true);
        session.Execute(new UpdateScriptSignatureCommand(modules, signature));
        Assert.True(session.IsDirty); Assert.False(session.CanUndo); Assert.False(session.ScriptsEnabled);
        session.Execute(new UpdateScriptModulesCommand(new ScriptModuleSet(modules.Modules, modules.EntryModules, scriptApiVersion: 2, bindings: [new("main", "actions"), new("helper", "actions")])));
        Assert.Null(session.Snapshot.ScriptModules.Signature);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new UpdateScriptSignatureCommand(modules, signature)));
    }

    [Fact]
    public void SignedPackageRoundTripsAndEditingCannotResurrectOldManifestSignature()
    {
        var storage = new TypedWorkbookStorage(); var path = PathFor("signed.xlsq"); var signed = Signed();
        storage.Save(path, Book(signed)); var opened = storage.Open(path);
        Assert.Equal(signed.Signature, opened.ScriptModules.Signature);
        Assert.Equal(ScriptSignatureStatus.UntrustedSelfSigned, ScriptSignatures.Verify(opened.ScriptModules).Status);
        using var session = new DocumentSession(opened); Assert.False(session.ScriptsEnabled);
        session.Execute(new UpdateScriptSourceCommand("edited")); storage.Save(path, session.Snapshot);
        var edited = storage.Open(path); Assert.Null(edited.ScriptModules.Signature);
        Assert.DoesNotContain("script-signatures", edited.FileState!.ManifestJson);
        Assert.DoesNotContain("scriptSignature", edited.FileState.ManifestJson);
    }

    [Fact]
    public void SignatureFeatureAndReferenceMustAgree()
    {
        var storage = new TypedWorkbookStorage(); var path = PathFor("broken.xlsq"); storage.Save(path, Book(Signed()));
        var budget = new FileReadBudget(CancellationToken.None);
        var parts = PackageIO.Unzip(File.ReadAllBytes(path), budget, 72, true);
        var manifest = JsonNode.Parse(parts["manifest.json"])!.AsObject();
        manifest["requiredFeatures"]!.AsArray().RemoveAt(3);
        parts["manifest.json"] = System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(manifest);
        File.WriteAllBytes(path, PackageIO.Zip(parts));
        Assert.Throws<InvalidDataException>(() => storage.Open(path));
    }

    [Fact]
    public void CorruptTrustFileFailsClosedAndNeverGetsSilentlyOverwritten()
    {
        var path = PathFor("trust.json"); File.WriteAllText(path, "broken");
        var store = new ScriptCertificateTrustStore(path);
        Assert.ThrowsAny<System.Text.Json.JsonException>(() => store.Verify(Signed()));
        Assert.ThrowsAny<System.Text.Json.JsonException>(() => store.Trust(Signed()));
        Assert.Equal("broken", File.ReadAllText(path));
    }

    [Fact]
    public void PrivateKeyStaysInPasswordProtectedPfxAndIsNotInWorkbook()
    {
        var pfx = _certificate.Export(X509ContentType.Pkcs12, "test-pass");
        using var imported = X509CertificateLoader.LoadPkcs12(pfx, "test-pass", X509KeyStorageFlags.EphemeralKeySet);
        var signature = ScriptSignatures.Sign(Modules(), imported);
        using var publicOnly = X509CertificateLoader.LoadCertificate(Convert.FromBase64String(signature.Certificate));
        Assert.False(publicOnly.HasPrivateKey);
        Assert.Throws<CryptographicException>(() => X509CertificateLoader.LoadPkcs12(pfx, "wrong", X509KeyStorageFlags.EphemeralKeySet));
    }

    public void Dispose() { _certificate.Dispose(); Directory.Delete(_folder, true); }
}
