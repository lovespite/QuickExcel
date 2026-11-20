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
    private ScriptSigningPreferencesStore SigningPreferences => new(Path.Combine(_userDataDirectory, "script-signing-preferences.json"));
    private static string CertificateDetails(ScriptSignatureVerification result) =>
        UiText.Format($"{result.Message}\n\n主体：{result.Subject ?? "—"}\n颁发者：{result.Issuer ?? "—"}\n有效期至：{result.NotAfter:yyyy-MM-dd HH:mm:ss}\nSHA-256 指纹：\n{result.Fingerprint ?? "—"}");
    private async Task<bool> CheckScriptTrustAsync()
    {
        RequireScriptsEnabled();
        var modules = Document.Snapshot.ScriptModules;
        if (modules.Signature is null) return true; // Explicit Apply is required for unsigned sources.
        DisableScripts(preserveNetworkConsent: true);
        ScriptSignatureVerification result; string? protectionFailure=null;
        try { result=await Task.Run(() => ScriptTrust.Verify(modules)); }
        catch(UserProtectionUnavailableException ex) {protectionFailure=ex.Message;result=await Task.Run(()=>ScriptSignatures.Verify(modules));}
        if (!ReferenceEquals(modules, Document.Snapshot.ScriptModules)) return false;
        if (result.CanAutoEnable) return true;
        if (result.Status == ScriptSignatureStatus.UntrustedSelfSigned)
        {
            var dialog = new FieldsDialog(UiText.Get("核对脚本签名证书")).Note(CertificateDetails(result))
                .Note(protectionFailure ?? UiText.Get("始终信任会允许此证书签署的其他工作簿自动启用脚本。请确认发布者和指纹。"))
                .Choice("trust", UiText.Get("信任选择"), protectionFailure is null ? [UiText.Get("取消，保持停用"), UiText.Get("仅本次允许"), UiText.Get("始终信任此证书")] : [UiText.Get("取消，保持停用"), UiText.Get("仅本次允许")]);
            if (!await dialog.ShowAsync(this) || !ReferenceEquals(modules, Document.Snapshot.ScriptModules)) return false;
            if (dialog.Index("trust") == 1) return true;
            if (dialog.Index("trust") == 2)
            {
                await Task.Run(() => ScriptTrust.Trust(modules));
                return ReferenceEquals(modules, Document.Snapshot.ScriptModules) && (await Task.Run(() => ScriptTrust.Verify(modules))).CanAutoEnable;
            }
            return false;
        }
        Error(new InvalidOperationException(result.Message + UiText.Get(" 请检查签名，或审阅源码后移除签名再手动运行。"))); return false;
    }
    private async Task ScriptSignatureDetailsAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        var modules = Document.Snapshot.ScriptModules;
        ScriptSignatureVerification result;
        try { result = await Task.Run(() => ScriptTrust.Verify(modules)); }
        catch (UserProtectionUnavailableException) { result = await Task.Run(() => ScriptSignatures.Verify(modules)); }
        var dialog = new FieldsDialog(UiText.Get("脚本签名详情")).Note(CertificateDetails(result)).Note(UiText.Get("签名保护全部源码与用途绑定，不包含单元格数据。"))
            .Choice("action", UiText.Get("操作"), modules.Signature is null ? [UiText.Get("关闭"), UiText.Get("签署脚本")] : [UiText.Get("关闭"), UiText.Get("签署脚本"), UiText.Get("移除签名并停用")]);
        if (!await dialog.ShowAsync(this)) return;
        if (dialog.Index("action") == 1) await SignScriptsAsync();
        else if (dialog.Index("action") == 2)
        { _workspace.Editor.Execute(new UpdateScriptSignatureCommand(modules, null)); DisableScripts(); Synchronize(); }
    }
    private async Task SignScriptsAsync()
    {
        RequireScriptsEnabled();
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        ScriptSigningPreferences remembered = new(); string? memoryError = null;
        try { remembered = await Task.Run(() => SigningPreferences.Read()); }
        catch (Exception ex) when (ex is IOException or CryptographicException or System.Text.Json.JsonException)
        { memoryError = UiText.Get("无法读取本机签名记忆，请重新选择私钥并输入密码：") + ex.Message; }
        var hasRememberedFile = remembered.RememberKeyFile;
        var choice = new FieldsDialog(UiText.Get("签署全部脚本")).Note(UiText.Get("修改源码或用途会移除签名并撤销授权；签名后请保存工作簿。私钥不会写入工作簿。"))
            .Choice("kind", UiText.Get("证书"), hasRememberedFile
                ? [UiText.Get("使用上次私钥文件"), UiText.Get("使用 PFX / P12 证书"), UiText.Get("创建自签证书")]
                : [UiText.Get("使用 PFX / P12 证书"), UiText.Get("创建自签证书")]);
        if (hasRememberedFile) choice.Note(remembered.KeyFilePath!);
        if (memoryError is not null) choice.Note(memoryError);
        if (!await choice.ShowAsync(this)) return;
        var create = choice.Index("kind") == (hasRememberedFile ? 2 : 1);
        LocalFileSelection? file = create ? await _files.SaveAsync(new(UiText.Get("保存签名证书与私钥"), [new(UiText.Get("PFX 证书"), ["*.pfx"])], "QuickExcel-signing.pfx", ".pfx")) :
            hasRememberedFile && choice.Index("kind") == 0 ? new(remembered.KeyFilePath!) :
            await _files.OpenAsync(new(UiText.Get("选择包含私钥的证书"), [new(UiText.Get("签名证书"), ["*.pfx", "*.p12"])]));
        if (file is null) return;
        var dialog = new FieldsDialog(create ? UiText.Get("创建自签证书") : UiText.Get("证书密码"));
        if (create) dialog.Text("name", UiText.Get("证书名称"), "QuickExcel Script Publisher");
        dialog.Text("password", create ? UiText.Get("PFX 密码（至少 8 个字符）") : UiText.Get("PFX 密码"),
            create ? "" : await Task.Run(() => remembered.PasswordForFile(file.Path)));
        ((TextBox)dialog.Input("password")).PasswordChar = '●';
        dialog.Check("rememberFile", UiText.Get("记住上次的私钥文件"), remembered.RememberKeyFile)
            .Check("rememberPassword", UiText.Get("记住密码"), remembered.RememberPassword)
            .Note(UiText.Get("仅保存在本机当前用户的受保护设置中，不写入 XLSQ。取消勾选并完成签名后会清除对应记忆。"));
        dialog.Validate = () => { if (create && dialog.String("password").Length < 8) throw new ArgumentException(UiText.Get("密码至少需要 8 个字符。")); };
        if (!await dialog.ShowAsync(this)) { ((TextBox)dialog.Input("password")).Text = ""; return; }
        var secret = dialog.String("password"); ((TextBox)dialog.Input("password")).Text = "";
        var rememberFile = dialog.Boolean("rememberFile"); var rememberPassword = dialog.Boolean("rememberPassword");
        var subject = create ? dialog.String("name") : ""; var document = Document; var modules = document.Snapshot.ScriptModules;
        var stamp = new CalculationStamp(document.Session.SessionId, document.Snapshot.Revision, 0, modules.Revision);
        try
        {
            _status.Text = UiText.Get("正在编译并签署全部脚本…");
            var signature = await Task.Run(async () =>
            {
                using var certificate = create ? ScriptSignatures.CreateSelfSigned(subject) : X509CertificateLoader.LoadPkcs12FromFile(file.Path, secret, X509KeyStorageFlags.EphemeralKeySet);
                var signature = await ScriptSignatures.SignAsync(modules, certificate, stamp);
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
            if (!ReferenceEquals(document, Document)) throw new InvalidOperationException(UiText.Get("工作簿已改变，请重新签名。"));
            _workspace.Editor.Execute(new UpdateScriptSignatureCommand(modules, signature)); DisableScripts(); Synchronize(); _status.Text = UiText.Get("已编译并签署全部脚本，请保存工作簿以写入编译缓存。");
            try { await Task.Run(() => SigningPreferences.Save(file.Path, secret, rememberFile, rememberPassword)); }
            catch (Exception ex) when (ex is IOException or CryptographicException)
            { _status.Text += UiText.Get(" 未能保存本机签名记忆：") + ex.Message; }
        }
        catch (QuickExcel.Runtime.ScriptCompilationException ex)
        {
            Error(new InvalidOperationException(UiText.Get("脚本编译未通过，未签名：") + "\n" +
                string.Join("\n", ex.Diagnostics.Take(20).Select(d => $"{d.Module}:{d.Line}:{d.Column} · {d.Message}")), ex));
        }
        finally { secret = ""; }
    }
    private async Task TrustedCertificatesAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var entries = await Task.Run(() => ScriptTrust.Read());
        if (entries.Count == 0) { _status.Text = UiText.Get("尚未记住任何脚本签名证书。"); return; }
        var dialog = new FieldsDialog(UiText.Get("受信任脚本证书")).Note(UiText.Get("移除后当前工作簿脚本会停用。系统证书信任由操作系统管理。"))
            .Choice("certificate", UiText.Get("证书及 SHA-256 指纹"), entries.Select(c => $"{c.Subject} · {c.NotAfter:yyyy-MM-dd}\n{c.Fingerprint}").ToArray())
            .Check("remove", UiText.Get("移除选中的证书"));
        if (!await dialog.ShowAsync(this) || !dialog.Boolean("remove")) return;
        await Task.Run(() => ScriptTrust.Remove(entries[dialog.Index("certificate")].Fingerprint)); DisableScripts();
    }
}
