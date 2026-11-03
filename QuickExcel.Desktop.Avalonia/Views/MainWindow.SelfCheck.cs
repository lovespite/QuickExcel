using System.Text.Json;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation;
using QuickExcel.Runtime;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    // Available in production for installation diagnosis. No injected dialogs,
    // failure RPCs, test workers, test certificates or synthetic browser edits.
    private async Task RunSelfCheckAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool passed, string message) { if (!passed) throw new InvalidDataException(message); checks.Add(message); }
        try
        {
            _ = new WebAssetCatalog(Path.Combine(AppContext.BaseDirectory, "Web")); Check(true, "Production resource hashes and mode verified");
            WebRuntimePrerequisite.Require(); Check(true, "Platform browser runtime preflight");
            ShowBook(); _workspace.New();
            var cell = new CellAddress(Document.Snapshot.ActiveSheetId, 0, 0);
            _workspace.Editor.Execute(new SetCellsCommand([new(cell, CellInput.Parse("安装检查 中文 123"))]));
            var file = Path.Combine(directory, "installation.xlsq");
            await _workspace.SaveAsync(new(file)); await _workspace.OpenAsync(new(file));
            Check(Document.GetCellInput(0, 0, 0) == "安装检查 中文 123", "Typed workbook save/reopen");
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "quickExcel.functions.register('INSTALLCHECK', value => value + 1);") }, ["main"], scriptApiVersion: 2)));
            await OpenScriptsAsync(); Check(_scriptPage!.Browser.IsReady && await CommitScriptsAsync(), "Production Monaco worker and unchanged draft confirmation");
            await CloseScriptWindowAsync();
            if (DesktopServices.SupportsScripts)
            {
                using var runtime = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create("quickExcel.functions.register('INSTALLCHECK', value => value + 1);"));
                Check(Convert.ToDouble(await Task.Run(() => runtime.Invoke("INSTALLCHECK", [41d]))) == 42, "Published native bridge in platform-isolated worker");
                var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.ts", "const value:number=42; quickExcel.functions.register('TSCHECK',()=>value);", ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 2);
                using var compiled = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default)));
                Check(Convert.ToDouble(await Task.Run(() => compiled.Invoke("TSCHECK", []))) == 42, "Published isolated TypeScript compiler and verified emitted ESM");
            }
            await OpenPrintPreviewAsync(); Check(_pdfPage!.Browser.IsReady && _printSession!.Job.Layout.Pages.Length > 0, "Production PDF.js worker and complete print-session PDF");
            await _printSession!.ExportAsync(Path.Combine(directory, "installation.pdf"));
            await ClosePrintPreviewAsync(); Check(true, "PDF export and browser/resource cleanup");
        }
        catch (Exception ex) { failure = ex; }
        finally { try { await CloseScriptWindowAsync(); await ClosePrintPreviewAsync(); } catch (Exception ex) { failure ??= ex; } }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString(), webViewRuntime = WebRuntimePrerequisite.Version,
            softwareRenderingRequested = Program.SoftwareRendering, operatingSystem = Environment.OSVersion.ToString(), framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
