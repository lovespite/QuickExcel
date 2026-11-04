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
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "host.functions.register('INSTALLCHECK', value => value + 1);") }, ["main"], scriptApiVersion: 3)));
            await OpenScriptsAsync(); Check(_scriptPage!.Browser.IsReady && await CommitScriptsAsync(), "Production Monaco worker and unchanged draft confirmation");
            await CloseScriptWindowAsync();
            if (DesktopServices.SupportsScripts)
            {
                using var runtime = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create("host.functions.register('INSTALLCHECK', value => value + 1);"));
                Check(Convert.ToDouble(await Task.Run(() => runtime.Invoke("INSTALLCHECK", [41d]))) == 42, "Published native bridge in platform-isolated worker");
                var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.ts", "const value:number=42; host.functions.register('TSCHECK',()=>value);", ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 3);
                using var compiled = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default)));
                Check(Convert.ToDouble(await Task.Run(() => compiled.Invoke("TSCHECK", []))) == 42, "Published isolated TypeScript compiler and verified emitted ESM");
                var action = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.ts", """
                    function main(workbook: QuickExcel.Workbook) {
                      const sheet=workbook.getActiveWorksheet(), range=sheet.getRange('B1:C3');
                      range.setValues([['Name','Value'],['b',2],['a',1]]);
                      range.getFormat().getFont().setBold(true);
                      range.getFormat().setHorizontalAlignment(QuickExcel.HorizontalAlignment.center);
                      const table=sheet.addTable(range,true); table.setName('InstallTable');
                      table.getSort().apply([{key:1}]);
                      table.getAutoFilter().apply(table.getRange(),1,{filterOn:QuickExcel.FilterOn.values,values:['1']});
                      if(!table.getAutoFilter().getIsDataFiltered()||table.getRangeBetweenHeaderAndTotal().getValues()[0][1]!==1)throw Error('P1 verification');
                    }
                    """, ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 3, bindings: [new("main", "actions")]);
                _workspace.Editor.Execute(new UpdateScriptModulesCommand(action));
                // Use the normal execution gate so the UI's automatic
                // recalculation cannot supersede this explicit action request.
                _calculationTimer.Stop();
                using var actionCancellation = new CancellationTokenSource();
                using var actionFeedback = BeginExecution("正在验证脚本 API…", actionCancellation);
                var actionDocument = Document;
                await Task.Run(() => { actionDocument.EnableScriptsAndCalculate(actionCancellation.Token); actionDocument.RunAction("main", actionCancellation.Token); }, actionCancellation.Token);
                Check(Document.Snapshot.ScriptModules.ScriptApiVersion == 3 && Document.GetCellValue(0, 1, 2) == "1" &&
                    Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(1, 2)].Bold,
                    "Published API v3 enums, formatting, table, sort and filter transaction");
                Check(Document.Undo(), "Published API v3 action undo");
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
