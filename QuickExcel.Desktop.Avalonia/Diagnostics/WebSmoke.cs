using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Security;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunWebSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? error = null; var watch = Stopwatch.StartNew();
        void Check(bool value, string description) { if (!value) throw new InvalidOperationException(description); checks.Add(description); File.AppendAllText(Path.Combine(directory, "progress.log"), description + "\n"); }
        async Task<JsonElement> Test(object payload) => (await _scriptPage!.Browser.RequestAsync("test", "testResult", payload, TimeSpan.FromSeconds(30))).Payload;
        async Task SelectModule(string module)
        {
            var picker = _modulePicker!.Items.Contains(module) ? _modulePicker : _scriptModuleGroups.First(g => g.Picker.Items.Contains(module)).Picker;
            picker.SelectedItem = module; await _operation;
            for (var retry = 0; retry < 100 && _working; retry++) await Task.Delay(10);
            if (_working) throw new TimeoutException("Module navigation did not release its editor hold.");
        }
        try
        {
            ShowBook();
            if (Program.WebModulesOnly) await CheckScriptNavigationAsync(Test, SelectModule, Check, directory);
            else
            {
            _workspace.New(); _files = new SmokeFiles { NextSave = Path.Combine(directory, "script-roundtrip.xlsq") };
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule>
            { ["main"] = new("scripts/main.js", "import { twice } from 'helpers';\nhost.functions.register('DOUBLE', twice);\n"), ["helpers"] = new("scripts/helpers.js", "export function twice(value) { return value * 2; }\n") }, ["main"], scriptApiVersion: 3)));
            await OpenScriptsAsync(); Check(_scriptPage!.Browser.IsReady && !_scriptPage.Drafts.Failed, "Actual native WebView loads offline Monaco and a real language worker");
            var inspected = await Test(new { action = "inspect" });
            Check(inspected.GetProperty("workers").GetInt32() > 0 && inspected.GetProperty("source").GetString()!.Contains("DOUBLE"), "Language worker and initial module source are available");
            Check(inspected.GetProperty("diagnostics").GetArrayLength() == 0, "Declared logical module imports and actual QuickExcel API type-check without DOM or Node libraries");
            await CheckColdJavaScriptDefinitionsAsync(Test, Check);
            var revision = Document.Snapshot.Revision;
            const string changed = "import { twice } from 'helpers';\nhost.functions.register('DOUBLE', value => twice(value) + 1);\n";
            await Test(new { action = "edit", module = "main", text = changed });
            Check(Document.Snapshot.Revision == revision, "Browser editing does not create per-keystroke workbook history");
            await SelectModule("helpers"); inspected = await Test(new { action = "inspect" });
            Check(inspected.GetProperty("selected").GetString() == "helpers" && Document.Snapshot.ScriptModules.Modules["main"].Source == changed,
                "Left file navigation flushes the current draft before switching the actual editor model");
            await SelectModule("main");
            Check(await SaveAsync(false) && Document.Snapshot.ScriptModules.Modules["main"].Source == changed, "Save waits for complete draft versions and commits current source");
            var reopened = _services.Storage.Open(Path.Combine(directory, "script-roundtrip.xlsq"));
            Check(reopened.ScriptModules.Modules["main"].Source == changed, "Saved XLSQ contains the acknowledged source, not the previous source");
            await SelectModule("helpers"); await SelectModule("main"); await Test(new { action = "undo" });
            Check(await CommitScriptsAsync() && Document.Snapshot.ScriptModules.Modules["main"].Source.Contains("'DOUBLE', twice"), "Switching Monaco models preserves each model's undo history");
            await Test(new { action = "edit", module = "main", text = "/** @type {number} */\nlet value = '中文';\ndocument.querySelector('body');\nprocess.exit(0);\n" });
            inspected = await Test(new { action = "inspect" });
            await File.WriteAllTextAsync(Path.Combine(directory, "language-diagnostics.json"), inspected.ToString());
            Check(inspected.GetProperty("diagnostics").EnumerateArray().Any(d => d.GetProperty("code").GetInt32() == 2322), "JavaScript semantic diagnostics report incorrect annotated types");
            Check(inspected.GetProperty("diagnostics").EnumerateArray().Any(d => d.GetProperty("text").GetString()!.Contains("document")), "DOM globals are absent from the workbook runtime language environment");
            await Test(new { action = "edit", module = "main", text = "host.functions." });
            inspected = await Test(new { action = "inspect" });
            Check(inspected.GetProperty("completions").EnumerateArray().Any(v => v.GetString() == "register"), "QuickExcel completion is provided by the language worker");
            await Test(new { action = "edit", module = "main", text = "host" });
            await Test(new { action = "assistance", offset = "host".Length, typeText = "." });
            for (var retry = 0; retry < 100; retry++)
            {
                inspected = await Test(new { action = "assistance" });
                if (inspected.GetProperty("suggestions").GetString()!.Contains("functions")) break;
                await Task.Delay(50);
            }
            Check(inspected.GetProperty("suggestions").GetString()!.Contains("functions") && inspected.GetProperty("suggestions").GetString()!.Contains("events"),
                "The host extension opens the actual completion list on typing a dot");
            // Each host message remains below 256 KiB; drive the native editor with bounded test edits.
            // A smaller multi-chunk Unicode draft exercises the same checksum and total-source budget path.
            var large = "//" + string.Concat(Enumerable.Repeat("中文😀", 6_000));
            await Test(new { action = "edit", module = "main", text = large });
            Check(await CommitScriptsAsync() && Document.Snapshot.ScriptModules.PrimarySource == large, "Multi-chunk Unicode drafts reassemble and validate before a document command");
            await Test(new { action = "edit", module = "main", text = changed }); Check(await CommitScriptsAsync(), "Restored valid function source commits");
            if (DesktopServices.SupportsScripts) { await ApplyScriptsAsync(); Check(Document.ScriptsEnabled, "Explicit application uses the existing isolated worker"); }
            else
            {
                var blocked = false; try { await ApplyScriptsAsync(); } catch (PlatformNotSupportedException) { blocked = true; }
                Check(blocked && !Document.ScriptsEnabled, "Unimplemented platform sandbox fails closed without in-process execution");
            }
            await Test(new { action = "edit", module = "main", text = changed + "// changed\n" });
            Check(!Document.ScriptsEnabled, "Typing immediately revokes prior script execution authorization");
            Check(await CommitScriptsAsync(), "Latest script source is confirmed before simulated page failure");
            var confirmed = Document.Snapshot.ScriptModules.PrimarySource;
            try { await Test(new { action = "crash" }); } catch (Exception) when (_scriptPage.Drafts.Failed) { }
            Check(_scriptPage.Drafts.Failed && !await SaveAsync(false), "Page failure blocks save instead of falling back to older text");
            Check(_scriptPage.Drafts.Models["main"].Text == confirmed, "Confirmed source survives failed browser communication");
            _dialogDriver = window => { window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, "确定")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); return Task.CompletedTask; };
            await RecoverScriptPageAsync();
            Check(!_scriptPage!.Drafts.Failed && await SaveAsync(false) && Document.Snapshot.ScriptModules.PrimarySource == confirmed, "Explicit recovery creates a new page generation and preserves confirmed Unicode/source");
            await CloseScriptWindowAsync(); Check(_scriptPage is null && _scriptWindow is null, "Closing waits for editor/model/worker disposal and detaches native WebView");
            await CheckScriptNavigationAsync(Test, SelectModule, Check, directory);
            // Run the migrated actions/role/signature flows through the actual worker and native dialog handlers.
            if (DesktopServices.SupportsScripts)
            {
            _workspace.New();
            const string actionSource = "function main(workbook) { workbook.getActiveWorksheet().getRange('A1').setValue(42); console.log('原子写入完成'); }";
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", actionSource) }, ["main"], scriptApiVersion: 3, bindings: [new("main", "actions")])));
            await OpenScriptsAsync(); await ApplyScriptsAsync(); await RunScriptActionAsync();
            Check(Document.GetCellValue(0, 0, 0) == "42" && Document.ScriptErrors.Any(e => e.Contains("原子写入完成")), "Migrated action picker executes an immutable worker action and displays logs");
            DisableScripts(); await History(false); Check(Document.GetCellInput(0, 0, 0) == "", "An action remains one undoable transaction after authorization is revoked");
            await History(true); Check(Document.GetCellValue(0, 0, 0) == "42", "Action redo restores its patch without executing the script again");
            var modules = Document.Snapshot.ScriptModules;
            using (var certificate = ScriptSignatures.CreateSelfSigned("Avalonia isolated smoke"))
                _workspace.Editor.Execute(new UpdateScriptSignatureCommand(modules, ScriptSignatures.Sign(modules, certificate)));
            _dialogDriver = window => { var dialog = (FieldsDialog)window; ((ComboBox)dialog.Input("trust")).SelectedIndex = 1; dialog.Close(true); return Task.CompletedTask; };
            await ApplyScriptsAsync(); Check(Document.ScriptsEnabled && ScriptTrust.Read().Count == 0, "One-time signed-script consent does not persist a trust pin");
            DisableScripts(); _dialogDriver = window => { var dialog = (FieldsDialog)window; ((ComboBox)dialog.Input("trust")).SelectedIndex = 2; dialog.Close(true); return Task.CompletedTask; };
            await ApplyScriptsAsync(); Check(Document.ScriptsEnabled && ScriptTrust.Read().Count == 1, "Persistent certificate consent uses isolated integrity-protected trust storage");
            _dialogDriver = window =>
            {
                var dialog = (FieldsDialog)window;
                ((TextBox)dialog.Input("id")).Text = "udf"; ((ComboBox)dialog.Input("role")).SelectedIndex = 2;
                window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, "确定")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return Task.CompletedTask;
            };
            await AddScriptModuleAsync(); Check(Document.Snapshot.ScriptModules.Signature is null && !Document.ScriptsEnabled &&
                Document.Snapshot.ScriptModules.Bindings.SequenceEqual([new("main", "actions"), new("udf", "formula-functions")]),
                "Creating a bound module preserves existing roles, removes the signature and revokes authorization");
            _dialogDriver = window => { var dialog = (FieldsDialog)window; ((CheckBox)dialog.Input("remove")).IsChecked = true; dialog.Close(true); return Task.CompletedTask; };
            await TrustedCertificatesAsync(); Check(ScriptTrust.Read().Count == 0, "Native trust manager removes only the selected isolated trust pin");
            await CloseScriptWindowAsync();
            }
            _workspace.New();
            _files = new SmokeFiles { NextSave = Path.Combine(directory, "typescript-roundtrip.xlsq") };
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule>
            {
                ["main"] = new("scripts/main.ts", "import {twice} from 'helpers'; host.functions.register('DOUBLE',()=>twice(5));", ScriptLanguage.TypeScript),
                ["helpers"] = new("scripts/helpers.ts", "export const twice=(n:number):number=>n*2;", ScriptLanguage.TypeScript)
            }, ["main"], scriptApiVersion: 3)));
            _workspace.Editor.Execute(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 0, 0), CellInput.Parse("=DOUBLE()"))]));
            await OpenScriptsAsync();
            inspected = await Test(new { action = "inspect" });
            Check(inspected.GetProperty("language").GetString() == "typescript", "Monaco creates a TypeScript model with the actual TS language worker");
            await CheckScriptIntelligenceAsync(Test, Check, directory);
            await CheckScriptDefinitionsAsync(Test, Check, directory);
            var beforeCompile = Document.Snapshot.ScriptModules;
            await Test(new { action = "edit", module = "main", text = "export const value:number='类型错误';" });
            Check(!await CommitScriptsAsync(validateTypeScript: true) && ReferenceEquals(beforeCompile, Document.Snapshot.ScriptModules) && _scriptPage!.Drafts.IsDirty,
                "Failed isolated TS checking preserves the draft and does not replace applied source");
            Check(_scriptOutputTabs!.IsVisible && _scriptOutputTabs.SelectedIndex == 1, "Compiler failures expand runtime diagnostics even when the output pane was collapsed");
            _scriptDiagnostics!.SelectedIndex = 0; await ActivateScriptDiagnosticAsync();
            Check(_scriptPage!.Drafts.IsDirty, "Compiler diagnostic navigation retains failed draft text");
            await Test(new { action = "edit", module = "main", text = "import {twice} from 'helpers'; enum E { A=6 }; host.functions.register('DOUBLE',()=>twice(E.A));" });
            await ApplyScriptsAsync();
            Check(Document.ScriptsEnabled && Document.GetCellValue(0, 0, 0) == "12", "Apply automatically checks TS then runs emitted ESM in a fresh isolated execution worker");
            Check(await SaveAsync(false), "TS original sources save without an external toolchain");
            var tsSaved = new QuickExcel.Storage.TypedWorkbookStorage().Open(Path.Combine(directory, "typescript-roundtrip.xlsq"));
            Check(tsSaved.ScriptModules.Modules["main"].Language == ScriptLanguage.TypeScript && tsSaved.ScriptModules.PrimarySource.Contains("enum E"),
                "Saved TS reopens with its original language and source, not emitted JavaScript");
            await CloseScriptWindowAsync();
            if (DesktopServices.SupportsScripts) await CheckScriptConsoleAsync(Check, directory);
            if (!Program.WebScriptsOnly)
            {
            _workspace.New();
            _workspace.Editor.Execute(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 0, 0), CellInput.Parse("QuickExcel 中文打印预览")), new(new(Document.Snapshot.ActiveSheetId, 110, 2), CellInput.Parse("Frozen PDF content"))]));
            await OpenPrintPreviewAsync();
            Check(_pdfPage!.Browser.IsReady && _printSession!.Job.Layout.Pages.Length >= 2, "PDF.js loads a fully generated and independently verified multi-page PDF through its private resource lease");
            var previewFile = _printSession!.PdfPath;
            File.Copy(previewFile, Path.Combine(directory, "rendered-preview.pdf"));
            var pdfState = (await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "inspect" })).Payload;
            await File.WriteAllTextAsync(Path.Combine(directory, "pdf-first-page.json"), pdfState.ToString());
            Check(pdfState.GetProperty("canvases").GetInt32() is > 0 and <= 8 && pdfState.GetProperty("canvasBytes").GetInt64() <= 64 * 1024 * 1024, "Actual PDF page canvases stay within eight pages and 64 MiB");
            Check(pdfState.GetProperty("toolbarColor").GetString() == "rgb(36, 62, 59)", "PDF.js toolbar retains readable foreground colors with the system dark preference");
            var textFixture = Path.Combine(directory, "actualtext-fixture.pdf");
            Diagnostics.PdfTextFixture.Write(textFixture);
            await using (var fixtureLease = await _pdfPage.Browser.Origin.AddPdfAsync(textFixture))
            {
                var extraction = (await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "extractFixture", url = fixtureLease.Uri.AbsoluteUri })).Payload;
                await File.WriteAllTextAsync(Path.Combine(directory, "actualtext-extraction.json"), extraction.ToString());
                Check(extraction.GetProperty("text").GetString() == Diagnostics.PdfTextFixture.Expected, "Independent PDF ActualText preserves distinct same-glyph characters, ligatures, nested/empty spans and unmarked text");
            }
            string copied = "";
            for (var retry = 0; retry < 50 && !copied.Contains("中文"); retry++)
            { await Task.Delay(50); copied = (await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "copyText" })).Payload.GetProperty("copied").GetString()!; }
            await File.WriteAllTextAsync(Path.Combine(directory, "copied-text.txt"), copied);
            Check(copied.Contains("中文"), "PDF.js native text-layer copy handler returns Chinese text (synthetic clipboard, no OS clipboard mutation)");
            await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "find", text = "中文" });
            for (var retry = 0; retry < 50 && pdfState.GetProperty("matches").GetInt32() == 0; retry++)
            { await Task.Delay(50); pdfState = (await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "inspect" })).Payload; }
            Check(pdfState.GetProperty("matches").GetInt32() > 0, "PDF.js find controller locates text in the immutable PDF");
            await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "zoom", value = "3" });
            await Task.Delay(150); await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "page", page = _printSession.Job.Layout.Pages.Length });
            await Task.Delay(150); pdfState = (await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "inspect" })).Payload;
            Check(pdfState.GetProperty("page").GetInt32() == _printSession.Job.Layout.Pages.Length && pdfState.GetProperty("canvases").GetInt32() <= 8 && pdfState.GetProperty("canvasBytes").GetInt64() <= 64 * 1024 * 1024, "Zoom and page navigation replace old canvases while preserving the fixed cache budget");
            var job = _printSession.Job;
            _workspace.Editor.Execute(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 0, 0), CellInput.Parse("New workbook content"))])); Synchronize();
            Check(ReferenceEquals(job, _printSession.Job) && _printStatus!.Text!.Contains("工作簿已变化"), "Workbook edits leave the old preview frozen and visibly marked until explicit rebuild");
            try { await _pdfPage.Browser.RequestAsync("test", "testResult", new { action = "crash" }); } catch (InvalidDataException) { }
            Check(!_pdfPage.Browser.IsReady && File.Exists(previewFile) && _printWindow!.GetVisualDescendants().OfType<Button>().Any(b => Equals(b.Content, "导出 PDF")), "Failed preview retains its completed PDF and an independent native export control");
            ((SmokeFiles)_files).NextSave = Path.Combine(directory, "preview-export.pdf"); await ExportPreviewPdfAsync();
            Check(File.Exists(((SmokeFiles)_files).NextSave), "Native PDF export uses the same frozen print session");
            await RebuildPrintPreviewAsync(); Check(_pdfPage!.Browser.IsReady && !File.Exists(previewFile), "Explicit rebuild replaces a failed PDF page and retires the previous file");
            previewFile = _printSession!.PdfPath;
            await ClosePrintPreviewAsync();
            Check(!File.Exists(previewFile) && _pdfPage is null && _printSession is null, "Preview close destroys PDF worker and retires resource lease before deleting its temporary PDF");
            _workspace.New();
            _workspace.Editor.Execute(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 0, 0), CellInput.Parse("=UNAVAILABLE_PRINT_FUNCTION(1)"))]));
            var rejected = false;
            try { await OpenPrintPreviewAsync(); } catch (InvalidOperationException ex) when (ex.Message.Contains("受阻")) { rejected = true; }
            Check(rejected && _printSession is null && _pdfPage is null, "Blocked formulas cannot be printed from their old or missing caches");
            await ClosePrintPreviewAsync();
            }
            _workspace.New();
            var boundary = Enumerable.Range(0, 64).ToDictionary(i => "m" + i, i => new ScriptModule($"scripts/m{i}.js", ("export const value = " + i + ";\n//").PadRight(16 * 1024, ' ')));
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(boundary, ["m0"], scriptApiVersion: 3)));
            await OpenScriptsAsync();
            Check(_scriptPage!.Drafts.Models.Count == 64 && await CommitScriptsAsync(), "Real Monaco initializes and confirms all 64 modules at the exact 1 MiB total-source boundary");
            await CloseScriptWindowAsync();
            _workspace.New();
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(ScriptModuleSet.Single("host.functions.register('LEGACY', value => value);", 1)));
            await OpenScriptsAsync();
            Check(await CommitScriptsAsync() && Document.Snapshot.ScriptModules.ScriptApiVersion == 1, "Monaco preserves API v1 sources and does not migrate their execution contract");
            await CloseScriptWindowAsync();
            _workspace.Editor.Execute(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 0, 0), CellInput.Parse("=SUM(1,2)"))]));
            var readOnlyFile = Path.Combine(directory, "preserved-array.xlsx"); _services.Storage.Save(readOnlyFile, Document.Snapshot.WithScriptModules(ScriptModuleSet.Single("")));
            using (var zip = ZipFile.Open(readOnlyFile, ZipArchiveMode.Update))
            {
                XDocument xml; using (var input = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()) xml = XDocument.Load(input);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var formula = xml.Descendants(ns + "f").Single(); formula.SetAttributeValue("t", "array"); formula.SetAttributeValue("ref", "A1:A2");
                zip.GetEntry("xl/worksheets/sheet1.xml")!.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
            }
            await _workspace.OpenAsync(new(readOnlyFile)); await OpenScriptsAsync();
            inspected = await Test(new { action = "inspect" }); revision = Document.Snapshot.Revision;
            Check(Document.Snapshot.FileState!.FormulaReadOnly && inspected.GetProperty("readOnly").GetBoolean() && await CommitScriptsAsync() && Document.Snapshot.Revision == revision,
                "Preserved-formula workbooks open Monaco read-only and unchanged flush creates no mutation");
            inspected = await Test(new { action = "assistance" });
            Check(inspected.GetProperty("suggestDisabled").GetBoolean(), "Read-only workbooks disable the completion insertion control");
            if (DesktopServices.SupportsScripts)
            {
                _consoleInput!.Text = "workbook.getWorksheets().length"; await RunConsoleAsync();
                Check(ConsoleVisibleText.Contains("=> 1") && Document.Snapshot.Revision == revision,
                    "Read-only preserved workbooks allow console workbook inspection without mutation");
                _consoleInput.Text = "workbook.getActiveWorksheet().getRange('C1').setValue(42)"; await RunConsoleAsync();
                Check(Document.Snapshot.Revision == revision && Document.GetCellInput(0, 0, 2) == "" && _scriptStatus!.Text!.Contains("只读"),
                    "Read-only console rejects writes and keeps original workbook contents");
            }
            await CloseScriptWindowAsync();
        }
        }
        catch (Exception ex) { error = ex; }
        finally { try { await CloseScriptWindowAsync(); await ClosePrintPreviewAsync(); } catch (Exception ex) { error ??= ex; } }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = error is null, checks, error = error?.ToString(), elapsedMs = watch.Elapsed.TotalMilliseconds,
            platform = Environment.OSVersion.ToString(), skipped = DesktopServices.SupportsScripts ? Array.Empty<string>() : ["Platform sandbox is not qualified; execution remains closed"] }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}
