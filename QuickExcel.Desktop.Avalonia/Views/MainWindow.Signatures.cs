using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Security;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private ScriptCertificateTrustStore ScriptTrust => new(Path.Combine(_userDataDirectory, "trusted-script-certificates.json"));
    private static string CertificateDetails(ScriptSignatureVerification result) =>
        $"{result.Message}\n\n主体：{result.Subject ?? "—"}\n颁发者：{result.Issuer ?? "—"}\n有效期至：{result.NotAfter:yyyy-MM-dd HH:mm:ss}\nSHA-256 指纹：\n{result.Fingerprint ?? "—"}";
    private async Task<bool> CheckScriptTrustAsync()
    {
        var modules = Document.Snapshot.ScriptModules;
        if (modules.Signature is null) return true; // Explicit Apply is required for unsigned sources.
        DisableScripts();
        ScriptSignatureVerification result; string? protectionFailure=null;
        try { result=await Task.Run(() => ScriptTrust.Verify(modules)); }
        catch(UserProtectionUnavailableException ex) {protectionFailure=ex.Message;result=await Task.Run(()=>ScriptSignatures.Verify(modules));}
        if (!ReferenceEquals(modules, Document.Snapshot.ScriptModules)) return false;
        if (result.CanAutoEnable) return true;
        if (result.Status == ScriptSignatureStatus.UntrustedSelfSigned)
        {
            var dialog = new FieldsDialog("核对脚本签名证书").Note(CertificateDetails(result))
                .Note(protectionFailure ?? "始终信任会允许此证书签署的其他工作簿自动启用脚本。请确认发布者和指纹。")
                .Choice("trust", "信任选择", protectionFailure is null ? ["取消，保持停用", "仅本次允许", "始终信任此证书"] : ["取消，保持停用", "仅本次允许"]);
            if (!await dialog.ShowAsync(this) || !ReferenceEquals(modules, Document.Snapshot.ScriptModules)) return false;
            if (dialog.Index("trust") == 1) return true;
            if (dialog.Index("trust") == 2)
            {
                await Task.Run(() => ScriptTrust.Trust(modules));
                return ReferenceEquals(modules, Document.Snapshot.ScriptModules) && (await Task.Run(() => ScriptTrust.Verify(modules))).CanAutoEnable;
            }
            return false;
        }
        Error(new InvalidOperationException(result.Message + " 请检查签名，或审阅源码后移除签名再手动运行。")); return false;
    }
    private async Task ScriptSignatureDetailsAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        var modules = Document.Snapshot.ScriptModules;
        ScriptSignatureVerification result;
        try { result = await Task.Run(() => ScriptTrust.Verify(modules)); }
        catch (UserProtectionUnavailableException) { result = await Task.Run(() => ScriptSignatures.Verify(modules)); }
        var dialog = new FieldsDialog("脚本签名详情").Note(CertificateDetails(result)).Note("签名保护全部源码与用途绑定，不包含单元格数据。")
            .Choice("action", "操作", modules.Signature is null ? ["关闭", "签署脚本…"] : ["关闭", "签署脚本…", "移除签名并停用"]);
        if (!await dialog.ShowAsync(this)) return;
        if (dialog.Index("action") == 1) await SignScriptsAsync();
        else if (dialog.Index("action") == 2)
        { _workspace.Editor.Execute(new UpdateScriptSignatureCommand(modules, null)); DisableScripts(); Synchronize(); }
    }
    private async Task SignScriptsAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        var choice = new FieldsDialog("签署全部脚本").Note("修改源码或用途会移除签名并撤销授权；签名后请保存工作簿。私钥不会写入工作簿。")
            .Choice("kind", "证书", ["使用 PFX / P12 证书", "创建自签证书"]);
        if (!await choice.ShowAsync(this)) return;
        var create = choice.Index("kind") == 1;
        LocalFileSelection? file = create ? await _files.SaveAsync(new("保存签名证书与私钥", [new("PFX 证书", ["*.pfx"])], "QuickExcel-signing.pfx", ".pfx")) :
            await _files.OpenAsync(new("选择包含私钥的证书", [new("签名证书", ["*.pfx", "*.p12"])]));
        if (file is null) return;
        var dialog = new FieldsDialog(create ? "创建自签证书" : "证书密码");
        if (create) dialog.Text("name", "证书名称", "QuickExcel Script Publisher");
        dialog.Text("password", create ? "PFX 密码（至少 8 个字符）" : "PFX 密码"); ((TextBox)dialog.Input("password")).PasswordChar = '●';
        dialog.Validate = () => { if (create && dialog.String("password").Length < 8) throw new ArgumentException("密码至少需要 8 个字符。"); };
        if (!await dialog.ShowAsync(this)) return;
        var secret = dialog.String("password"); ((TextBox)dialog.Input("password")).Text = "";
        var subject = create ? dialog.String("name") : ""; var modules = Document.Snapshot.ScriptModules;
        try
        {
            var signature = await Task.Run(() =>
            {
                using var certificate = create ? ScriptSignatures.CreateSelfSigned(subject) : X509CertificateLoader.LoadPkcs12FromFile(file.Path, secret, X509KeyStorageFlags.EphemeralKeySet);
                var signature = ScriptSignatures.Sign(modules, certificate);
                if (create)
                {
                    var bytes = certificate.Export(X509ContentType.Pkcs12, secret);
                    var temporary = Path.Combine(Path.GetDirectoryName(file.Path)!, ".quickexcel-cert-" + Guid.NewGuid().ToString("N") + ".tmp");
                    try
                    {
                        using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { stream.Write(bytes); stream.Flush(true); }
                        File.Move(temporary, file.Path, true);
                    }
                    finally { CryptographicOperations.ZeroMemory(bytes); if (File.Exists(temporary)) File.Delete(temporary); }
                }
                return signature;
            });
            _workspace.Editor.Execute(new UpdateScriptSignatureCommand(modules, signature)); DisableScripts(); Synchronize(); _status.Text = "已签署全部脚本，请保存工作簿。";
        }
        finally { secret = ""; }
    }
    private async Task TrustedCertificatesAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var entries = await Task.Run(() => ScriptTrust.Read());
        if (entries.Count == 0) { _status.Text = "尚未记住任何脚本签名证书。"; return; }
        var dialog = new FieldsDialog("受信任脚本证书").Note("移除后当前工作簿脚本会停用。系统证书信任由操作系统管理。")
            .Choice("certificate", "证书及 SHA-256 指纹", entries.Select(c => $"{c.Subject} · {c.NotAfter:yyyy-MM-dd}\n{c.Fingerprint}").ToArray())
            .Check("remove", "移除选中的证书");
        if (!await dialog.ShowAsync(this) || !dialog.Boolean("remove")) return;
        await Task.Run(() => ScriptTrust.Remove(entries[dialog.Index("certificate")].Fingerprint)); DisableScripts();
    }
}
