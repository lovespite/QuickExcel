using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Security;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private ScriptCertificateTrustStore _scriptTrust = new(Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel", "trusted-script-certificates.json"));
    private ScriptModuleSet? _verifiedScriptModules;
    private ScriptSignatureVerification? _signatureVerification;
    private enum CertificateDecision { Cancel, Once, Always }
    private CertificateDecision? _smokeCertificateDecision;
    private int _certificatePromptCount;

    private void RefreshSignatureState()
    {
        if (SignatureState is null || _document is null) return;
        var modules = _document.Snapshot.ScriptModules;
        SignatureState.Text = modules.Signature is null ? "未签名" : ReferenceEquals(modules, _verifiedScriptModules)
            ? _signatureVerification?.Message ?? "已签名，待验证" : "已签名，待验证";
    }

    private async Task<ScriptSignatureVerification?> VerifyScriptSignatureAsync()
    {
        var modules = _document.Snapshot.ScriptModules;
        ScriptSignatureVerification? result = null;
        if (!await WorkAsync("验证脚本签名…", () => result = _scriptTrust.Verify(modules))) return null;
        if (!ReferenceEquals(modules, _document.Snapshot.ScriptModules)) return null;
        _verifiedScriptModules = modules; _signatureVerification = result;
        RefreshSignatureState();
        return result;
    }

    private async Task<bool> CheckScriptTrustAsync(bool automatic)
    {
        if (_document.Snapshot.ScriptModules.Signature is null) return !automatic;
        _document.DisableScripts(); ActionPicker.ItemsSource = null; ShowDiagnostics();
        var result = await VerifyScriptSignatureAsync();
        if (result is null) return false;
        if (result.CanAutoEnable) { AppendLog(result.Message); return true; }
        if (result.Status == ScriptSignatureStatus.UntrustedSelfSigned)
        {
            var modules = _document.Snapshot.ScriptModules;
            var decision = AskCertificateTrust(result);
            if (decision == CertificateDecision.Once) { AppendLog("本次允许此工作簿的已签名脚本。"); return true; }
            if (decision == CertificateDecision.Always)
            {
                if (!await WorkAsync("保存证书信任…", () => _scriptTrust.Trust(modules))) return false;
                return (await VerifyScriptSignatureAsync())?.CanAutoEnable == true;
            }
        }
        _document.DisableScripts(); ActionPicker.ItemsSource = null;
        AppendLog(result.Message + " 脚本保持停用。");
        if (!automatic && (result.Status is ScriptSignatureStatus.Invalid or ScriptSignatureStatus.Untrusted) && !_smokeMode)
            MessageBox.Show(this, result.Message + "\n请在签名详情中检查证书，或审阅源码后移除签名再手动运行。", "未允许脚本", MessageBoxButton.OK, MessageBoxImage.Warning);
        ShowDiagnostics(); UpdateChrome();
        return false;
    }

    private CertificateDecision AskCertificateTrust(ScriptSignatureVerification result)
    {
        _certificatePromptCount++;
        if (_smokeMode && _smokeSignatureFolder is null) return _smokeCertificateDecision ?? CertificateDecision.Cancel;
        var panel = new StackPanel { Margin = new(24) };
        panel.Children.Add(new TextBlock { Text = "此工作簿使用自签证书。请核对证书指纹后选择是否信任。\n始终信任后，此证书签署的其他工作簿也会在打开时自动启用脚本。", TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(CertificateDetails(result));
        var buttons = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 18, 0, 0) };
        panel.Children.Add(buttons);
        var dialog = ToolDialog("信任脚本签名证书", panel, 650, 430);
        dialog.MinWidth = 650; dialog.MinHeight = 430;
        var decision = CertificateDecision.Cancel;
        foreach (var (label, choice) in new[] { ("取消", CertificateDecision.Cancel), ("仅本次允许", CertificateDecision.Once), ("始终信任此证书", CertificateDecision.Always) })
        {
            var button = new Button { Content = label, Tag = choice, Margin = new(6, 0, 0, 0), IsCancel = choice == CertificateDecision.Cancel, IsDefault = choice == CertificateDecision.Cancel };
            button.Click += (_, _) => { decision = choice; dialog.Close(); }; buttons.Children.Add(button);
        }
        if (_smokeMode && _smokeSignatureFolder is not null) ExerciseCertificateDialog(dialog, buttons);
        dialog.ShowDialog();
        return decision;
    }

    private static TextBox CertificateDetails(ScriptSignatureVerification result) => new()
    {
        Text = $"{result.Message}\n\n主体：{result.Subject ?? "—"}\n颁发者：{result.Issuer ?? "—"}\n有效期至：{result.NotAfter:yyyy-MM-dd HH:mm:ss}\nSHA-256 指纹：\n{result.Fingerprint ?? "—"}",
        IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
        Height = 205, Margin = new(0, 16, 0, 0), BorderThickness = new(0)
    };

    private async void SignatureDetails_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var result = await VerifyScriptSignatureAsync();
        if (result is null) return;
        var panel = new StackPanel { Margin = new(24) };
        panel.Children.Add(CertificateDetails(result));
        panel.Children.Add(new TextBlock { Text = "签名保护全部脚本及用途绑定；单元格数据不在签名范围内。", TextWrapping = TextWrapping.Wrap, Margin = new(0, 12, 0, 12) });
        var remove = new Button { Content = "移除签名并停用脚本", IsEnabled = _document.Snapshot.ScriptModules.Signature is not null, HorizontalAlignment = HorizontalAlignment.Right };
        panel.Children.Add(remove);
        var dialog = ToolDialog("脚本签名详情", panel, 650, 410);
        remove.Click += (_, _) =>
        {
            var modules = _document.Snapshot.ScriptModules;
            _editor.Execute(new UpdateScriptSignatureCommand(modules, null));
            ActionPicker.ItemsSource = null; RefreshSignatureState(); ShowDiagnostics(); UpdateChrome(); dialog.Close();
        };
        dialog.ShowDialog();
    }

    private void SignScript_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var panel = new StackPanel { Margin = new(24) };
        panel.Children.Add(new TextBlock { Text = "使用带私钥的代码签名证书签署全部脚本。\n签名后请保存工作簿。修改脚本将移除签名并停用脚本。", TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 20) });
        var import = new Button { Content = "使用 PFX / P12 证书…", Margin = new(0, 0, 0, 12) };
        var create = new Button { Content = "创建自签证书并签名…" };
        panel.Children.Add(import); panel.Children.Add(create);
        var dialog = ToolDialog("签署工作簿脚本", panel, 480, 280);
        import.Click += async (_, _) => { dialog.Close(); await SignWithCertificateAsync(false); };
        create.Click += async (_, _) => { dialog.Close(); await SignWithCertificateAsync(true); };
        dialog.ShowDialog();
    }

    private async Task SignWithCertificateAsync(bool create)
    {
        if (_busy || !CommitEditors()) return;
        string path;
        if (create)
        {
            var save = new SaveFileDialog { Filter = "签名证书和私钥 (*.pfx)|*.pfx", DefaultExt = ".pfx", AddExtension = true, FileName = "QuickExcel-signing.pfx", Title = "保存自签证书和私钥" };
            if (save.ShowDialog(this) != true) return;
            path = save.FileName;
        }
        else
        {
            var open = new OpenFileDialog { Filter = "签名证书 (*.pfx;*.p12)|*.pfx;*.p12", Title = "选择含私钥的代码签名证书" };
            if (open.ShowDialog(this) != true) return;
            path = open.FileName;
        }
        var panel = new StackPanel { Margin = new(24) };
        var name = new TextBox { Text = "QuickExcel Script Publisher", Margin = new(0, 8, 0, 16) };
        if (create) { panel.Children.Add(new TextBlock { Text = "证书名称" }); panel.Children.Add(name); }
        panel.Children.Add(new TextBlock { Text = create ? "设置 PFX 密码（至少 8 个字符；请妥善保管私钥文件）" : "PFX 密码" });
        var password = new PasswordBox { Margin = new(0, 8, 0, 16) }; panel.Children.Add(password);
        var sign = new Button { Content = "签名", IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right }; panel.Children.Add(sign);
        var dialog = ToolDialog(create ? "创建自签证书" : "证书密码", panel, 530, create ? 340 : 260);
        sign.Click += (_, _) =>
        {
            if (create && password.Password.Length < 8) { password.Focus(); return; }
            dialog.DialogResult = true;
        };
        if (dialog.ShowDialog() != true) return;
        var secret = password.Password; password.Clear();
        var subject = name.Text;
        var modules = _document.Snapshot.ScriptModules;
        ScriptSignature? signature = null;
        try
        {
            if (!await WorkAsync("签署脚本…", () =>
            {
                using var certificate = create ? ScriptSignatures.CreateSelfSigned(subject)
                    : X509CertificateLoader.LoadPkcs12FromFile(path, secret, X509KeyStorageFlags.EphemeralKeySet);
                signature = ScriptSignatures.Sign(modules, certificate);
                if (create)
                {
                    var pfx = certificate.Export(X509ContentType.Pkcs12, secret);
                    try { File.WriteAllBytes(path, pfx); }
                    finally { CryptographicOperations.ZeroMemory(pfx); }
                }
            })) return;
            _editor.Execute(new UpdateScriptSignatureCommand(modules, signature));
            ActionPicker.ItemsSource = null;
            await VerifyScriptSignatureAsync();
            AppendLog("已签署全部脚本，请保存工作簿。" + (create ? " 自签证书已保存到所选 PFX 文件。" : ""));
            ShowDiagnostics(); UpdateChrome();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { secret = ""; }
    }

    private void TrustedCertificates_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var panel = new DockPanel { Margin = new(24) };
            var help = new TextBlock { Text = "QuickExcel 记住的自签证书。系统证书信任由 Windows 管理。\n移除后当前工作簿脚本会停用，下次打开时重新询问信任。", TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 12) };
            DockPanel.SetDock(help, Dock.Top); panel.Children.Add(help);
            var remove = new Button { Content = "移除选中证书", HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 12, 0, 0) };
            DockPanel.SetDock(remove, Dock.Bottom); panel.Children.Add(remove);
            var list = new ListBox(); panel.Children.Add(list);
            void Reload() => list.ItemsSource = _scriptTrust.Read().Select(c => new ListBoxItem { Content = $"{c.Subject}\n有效期至 {c.NotAfter:yyyy-MM-dd}\n{c.Fingerprint}", Tag = c }).ToArray();
            Reload();
            remove.Click += (_, _) =>
            {
                if (list.SelectedItem is not ListBoxItem { Tag: TrustedScriptCertificate certificate }) return;
                try
                {
                    _scriptTrust.Remove(certificate.Fingerprint); _document.DisableScripts(); ActionPicker.ItemsSource = null;
                    _verifiedScriptModules = null; RefreshSignatureState(); ShowDiagnostics(); UpdateChrome(); Reload();
                }
                catch (Exception ex) { ReportEditFailure(ex); }
            };
            ToolDialog("受信任的脚本证书", panel, 720, 460).ShowDialog();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
