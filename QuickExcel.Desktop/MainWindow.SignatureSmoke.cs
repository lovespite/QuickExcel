using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Security;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private string? _smokeSignatureFolder;
    private Exception? _signatureDialogFailure;

    private void ExerciseCertificateDialog(Window dialog, WrapPanel buttons)
    {
        dialog.ContentRendered += (_, _) => dialog.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() =>
        {
            try
            {
                dialog.UpdateLayout();
                var controls = buttons.Children.OfType<Button>().ToArray();
                if (controls.Length != 3 || !controls.Single(b => (CertificateDecision)b.Tag == CertificateDecision.Cancel).IsDefault ||
                    controls.Where(b => (CertificateDecision)b.Tag != CertificateDecision.Cancel).Any(b => b.IsDefault))
                    throw new Exception("Trust dialog must default to Cancel.");
                foreach (var button in controls)
                {
                    var bounds = button.TransformToAncestor(dialog).TransformBounds(new Rect(button.RenderSize));
                    if (bounds.Right > dialog.ActualWidth || bounds.Bottom > dialog.ActualHeight)
                        throw new Exception("Trust dialog button is clipped.");
                }
                var dpi = VisualTreeHelper.GetDpi(dialog);
                var image = new RenderTargetBitmap((int)Math.Ceiling(dialog.ActualWidth * dpi.DpiScaleX),
                    (int)Math.Ceiling(dialog.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
                image.Render(dialog);
                var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(image));
                using (var file = File.Create(Path.Combine(_smokeSignatureFolder!, "trust-dialog.png"))) encoder.Save(file);
                var choice = _smokeCertificateDecision ?? CertificateDecision.Cancel;
                controls.Single(b => (CertificateDecision)b.Tag == choice).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
            catch (Exception ex) { _signatureDialogFailure = ex; dialog.Close(); }
        }));
    }

    private async Task RunSignatureSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        _smokeSignatureFolder = folder;
        var assertions = new List<string>();
        void Check(bool condition, string label)
        { if (!condition) throw new Exception(label + "\n" + LogText.Text); assertions.Add(label); }
        using var certificate = ScriptSignatures.CreateSelfSigned("QuickExcel signature smoke");
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", "quickExcel.functions.register('DOUBLE', x => x * 2); quickExcel.actions.register('Change', ({snapshot,patch}) => patch.setValue(snapshot.sheets[0].id,'C1',99));")
        }, ["main"], scriptApiVersion: 2, bindings: [new("main", "formula-functions"), new("main", "actions")]);
        modules = modules.WithSignature(ScriptSignatures.Sign(modules, certificate));
        var signedPath = Path.Combine(folder, "signed.xlsq");
        using (var seed = WorkbookDocument.CreateBlank(_services))
        {
            seed.SetCell(0, 0, 0, "21"); seed.SetCell(0, 0, 1, "=DOUBLE(A1)");
            var snapshot = seed.Snapshot.WithScriptModules(modules);
            _services.Storage.Save(signedPath, snapshot);
            _services.Storage.Save(Path.Combine(folder, "unsigned.xlsq"), snapshot.WithScriptModules(modules.WithSignature(null)));
            _services.Storage.Save(Path.Combine(folder, "tampered.xlsq"), snapshot.WithScriptModules(modules.WithSignature(modules.Signature! with { Value = Convert.ToBase64String(new byte[384]) })));
        }
        await OpenPathAsync(Path.Combine(folder, "unsigned.xlsq"));
        Check(!_document.ScriptsEnabled && _certificatePromptCount == 0, "Unsigned open stays disabled without prompt");
        _smokeCertificateDecision = CertificateDecision.Cancel;
        await OpenPathAsync(signedPath);
        Check(!_document.ScriptsEnabled && _certificatePromptCount == 1, "Declining self-signed certificate keeps scripts disabled");
        _smokeCertificateDecision = CertificateDecision.Once;
        await OpenPathAsync(signedPath);
        Check(_document.ScriptsEnabled && _document.GetCellValue(0, 0, 1) == "42", "Trust once runs signed formula through actual worker");
        Check(ActionPicker.Items.Count == 1 && _document.GetCellValue(0, 0, 2) == "", "Auto-enable discovers actions without invoking their bodies");
        Check(_scriptTrust.Read().Count == 0, "Trust once is not persisted");
        _smokeCertificateDecision = CertificateDecision.Cancel;
        await OpenPathAsync(signedPath);
        Check(!_document.ScriptsEnabled && _certificatePromptCount == 3, "Trust once is lost on reopen");
        _smokeCertificateDecision = CertificateDecision.Always;
        await OpenPathAsync(signedPath);
        Check(_document.ScriptsEnabled && _scriptTrust.Read().Count == 1, "Always trust persists after successful verification");
        _scriptTrust = new(Path.Combine(folder, "trusted-script-certificates.json"));
        _smokeCertificateDecision = CertificateDecision.Cancel;
        var prompts = _certificatePromptCount;
        await OpenPathAsync(signedPath);
        Check(_document.ScriptsEnabled && _certificatePromptCount == prompts, "Remembered certificate auto-enables without dialog");
        Check(!_document.IsDirty, "Verification and automatic calculation do not dirty the workbook");
        var trustPath = Path.Combine(folder, "trusted-script-certificates.json");
        var trustData = File.ReadAllText(trustPath);
        File.WriteAllText(trustPath, "broken");
        await RunScriptsAsync();
        Check(!_document.ScriptsEnabled, "A failed trust read revokes a previously enabled session");
        File.WriteAllText(trustPath, trustData);
        _editor.SetCell(0, 0, 0, "22");
        Check(_document.Snapshot.ScriptModules.Signature is not null, "Cell edits preserve script signature");
        await RunScriptsAsync();
        Check(_document.GetCellValue(0, 0, 1) == "44", "Signed formulas use current cell data");
        ScriptEditor.Text += "\n// edited";
        Check(!_document.ScriptsEnabled && _document.Snapshot.ScriptModules.Signature is null && SignatureState.Text == "未签名", "Editor immediately removes signature and authorization");
        await OpenPathAsync(Path.Combine(folder, "tampered.xlsq"));
        Check(!_document.ScriptsEnabled && _signatureVerification?.Status == ScriptSignatureStatus.Invalid && _certificatePromptCount == prompts,
            "Tampered signature cannot prompt for trust or auto-enable");
        await RunScriptsAsync();
        Check(!_document.ScriptsEnabled, "Manual Run also rejects an invalid signature");
        await OpenPathAsync(signedPath, allowAutomaticScripts: false);
        Check(!_document.ScriptsEnabled, "Recovery open suppresses automatic execution even for trusted signer");
        _scriptTrust.Remove(_scriptTrust.Read().Single().Fingerprint);
        await OpenPathAsync(signedPath);
        Check(!_document.ScriptsEnabled && _certificatePromptCount == prompts + 1, "Removing trust restores the prompt on reopen");
        Check(_signatureDialogFailure is null, "Actual trust dialog buttons and safe default pass: " + _signatureDialogFailure?.Message);
        ShowWorkbookPage(true); UpdateLayout(); CaptureGridEvidence(Path.Combine(folder, "script-signatures.png"));
        File.WriteAllText(Path.Combine(folder, "signature-smoke-result.txt"), "PASS: " + assertions.Count + " signature assertions\n" + string.Join("\n", assertions));
    }
}
