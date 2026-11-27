using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ColorTextBlock.Avalonia;
using Markdown.Avalonia;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using QuickExcel.Calculation;

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
            Check(!_book.IsVisible && _workbookMenuItems.All(item => !item.IsEnabled) &&
                _menu.Items.OfType<global::Avalonia.Controls.MenuItem>().Skip(1).All(item => !item.IsEnabled),
                "Home disables workbook commands before a document is opened or created");
            _ = new WebAssetCatalog(Path.Combine(AppContext.BaseDirectory, "Web")); Check(true, "Production resource hashes and mode verified");
            var sampleCatalog=new WorkbookSampleCatalog(Path.Combine(AppContext.BaseDirectory,"samples"));
            Check(sampleCatalog.Samples.Count==9&&sampleCatalog.Samples.All(sample=>
            {
                var snapshot=sampleCatalog.Read(sample.Id);
                return snapshot.ScriptModules.ScriptApiVersion==4&&snapshot.ScriptModules.Signature is null&&
                    snapshot.Sheets.Length==4&&snapshot.ScriptModules.Modules.ContainsKey("example");
            }),"Published sample catalog and all nine workbook payloads are complete and hash verified");
            var samplePicker=new SampleWorkbookDialog(sampleCatalog.Samples) { ShowActivated=false };
            try
            {
                samplePicker.Show(this);
                await Dispatcher.UIThread.InvokeAsync(()=>{},DispatcherPriority.Background);
                Check(samplePicker.GetVisualDescendants().OfType<ListBox>().Single(list=>list.Name=="SampleWorkbooks").Items.Count==9&&
                    _start.GetVisualDescendants().OfType<Button>().Any(button=>button.Name=="OpenSampleWorkbooks"&&button.IsEnabled),
                    "Published home button and sample picker expose nine examples");
            }
            finally { samplePicker.Close(); }
            // Exercise the actual license viewer in the trimmed package, including
            // its compiled styles and rendered Markdown text rather than only the input.
            var settings = new SettingsWindow(_settings, _recentFiles, SettingsSection.About) { ShowActivated = false };
            try
            {
                settings.Show(this);
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                var markdown = settings.GetVisualDescendants().OfType<MarkdownScrollViewer>().Single();
                Check(markdown.IsEffectivelyVisible && markdown.Bounds.Width > 0 && markdown.Bounds.Height > 0 &&
                    markdown.Markdown?.StartsWith("# Third-party components and licenses", StringComparison.Ordinal) == true &&
                    markdown.GetVisualDescendants().OfType<CTextBlock>().Any(text => text.Bounds.Width > 0 && text.Bounds.Height > 0 &&
                        text.Text.Contains("Third-party components and licenses", StringComparison.Ordinal)),
                    "Published Settings renders the bundled third-party Markdown notice");
                var selector = settings.GetVisualDescendants().OfType<ComboBox>().Single(control => control.Name == "SettingsLicenseSelector");
                selector.SelectedItem = selector.Items.OfType<string>().Single(path => path.EndsWith("OFFICE-SCRIPTS-NOTICE.md", StringComparison.Ordinal));
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Check(markdown.IsEffectivelyVisible && markdown.GetVisualDescendants().OfType<CTextBlock>()
                    .Any(text => text.Bounds.Width > 0 && text.Bounds.Height > 0 && text.Text.Contains("CC BY 4.0", StringComparison.Ordinal)),
                    "Published Settings renders the Office Scripts attribution offline");
            }
            finally { settings.Close(); }
            WebRuntimePrerequisite.Require(); Check(true, "Platform browser runtime preflight");
            ShowBook(); _workspace.New();
            var cell = new CellAddress(Document.Snapshot.ActiveSheetId, 0, 0);
            _workspace.Editor.Execute(new SetCellsCommand([new(cell, CellInput.Parse("安装检查 中文 123"))]));
            var file = Path.Combine(directory, "installation.xlsq");
            await _workspace.SaveAsync(new(file)); await _workspace.OpenAsync(new(file));
            Check(Document.GetCellInput(0, 0, 0) == "安装检查 中文 123", "Typed workbook save/reopen");
            var database = new DatabaseConnectionOptions(DatabaseProvider.Sqlite, Path.Combine(directory, "installation.sqlite"));
            var databaseTable = new DatabaseTable("安装检查");
            var databaseRows = await Task.Run(() => DatabaseTransfer.ExportTable(database, databaseTable,
                Document.Snapshot, Document.Snapshot.ActiveSheetId, new CellRectangle(0, 0, 0, 0), false));
            var databaseRead = await Task.Run(() => DatabaseTransfer.ImportTable(database, databaseTable, false));
            Check(databaseRows == 1 && databaseRead.Workbook.Sheets[0].GetCell(0, 0)?.Value.TextValue == "安装检查 中文 123",
                "Published SQLite native provider creates, writes and reads Unicode values");
            using (var metadata = typeof(System.Data.Odbc.OdbcConnection).Assembly.GetManifestResourceStream("System.Data.Odbc.OdbcMetaData.xml")
                ?? throw new InvalidDataException("Missing ODBC provider metadata"))
            using (var schema = new System.Data.DataSet())
            {
                schema.ReadXml(metadata);
                Check(schema.Tables["DataSourceInformation"]?.Columns["GroupByBehavior"]?.DataType == typeof(System.Data.Common.GroupByBehavior) &&
                    schema.Tables["DataSourceInformation"]?.Columns["IdentifierCase"]?.DataType == typeof(System.Data.Common.IdentifierCase) &&
                    schema.Tables["DataSourceInformation"]?.Columns["SupportedJoinOperators"]?.DataType == typeof(System.Data.Common.SupportedJoinOperators),
                    "Published ODBC metadata resolves all facade-qualified schema enums without requiring a driver");
            }
            var exclusive = false;
            try { using var external = File.OpenRead(file); }
            catch (IOException) { exclusive = true; }
            Check(exclusive, "Opened XLSQ remains exclusive to its owning workspace");
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "host.functions.register('INSTALLCHECK', value => value + 1);") }, ["main"], scriptApiVersion: 3)));
            await OpenScriptsAsync(); Check(_scriptPage!.Browser.IsReady && await CommitScriptsAsync(), "Production Monaco worker and unchanged draft confirmation");
            await CloseScriptWindowAsync();
            if (DesktopServices.SupportsScripts)
            {
                ScriptExecutionPolicy.Current.RequireEnabled();
                using var runtime = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create("host.functions.register('INSTALLCHECK', value => value + 1);"));
                Check(Convert.ToDouble(await Task.Run(() => runtime.Invoke("INSTALLCHECK", [41d]))) == 42, "Published native bridge in platform-isolated worker");
                var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.ts", "const value:number=42; host.functions.register('TSCHECK',()=>value);", ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion: 3);
                using var compiled = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(modules, new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default)));
                Check(Convert.ToDouble(await Task.Run(() => compiled.Invoke("TSCHECK", []))) == 42, "Published isolated TypeScript compiler and verified emitted ESM");
                var libraries = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.ts", """
                    import {sumBy} from 'lodash'; import Decimal from 'decimal.js'; import dayjs from 'dayjs';
                    import Papa from 'papaparse'; import {mean} from 'simple-statistics'; import {z} from 'zod';
                    import Fuse from 'fuse.js'; import {from,op} from 'arquero'; import {add} from 'mathjs';
                    export function CHECK():number {
                      const rows=Papa.parse<{name:string,n:number}>('name,n\napple,20\nbanana,22',{header:true,dynamicTyping:true}).data;
                      if(sumBy(rows,'n')!==42 || mean(rows.map(r=>r.n))!==21)throw Error('collections');
                      if(new Fuse(rows,{keys:['name']}).search('aple')[0].item.n!==20)throw Error('search');
                      if(dayjs.duration({minutes:42}).asMinutes()!==42 || Number(add(20,22))!==42)throw Error('date/math');
                      if(!z.url().safeParse('https://example.com').success || !z.file().safeParse(new File(['x'],'x')).success)throw Error('data globals');
                      return z.number().parse(new Decimal(from(rows).rollup({n:op.sum('n')}).get('n',0) as number).toNumber());
                    }
                    """, ScriptLanguage.TypeScript) }, ["main"], scriptApiVersion:4);
                using var bundled = await Task.Run(() => new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(libraries, new(Guid.NewGuid(),0,0,0),TimeSpan.FromSeconds(5),default)));
                Check(Convert.ToDouble(await Task.Run(() => bundled.Invoke("JS.main.CHECK", []))) == 42,
                    "Published nine built-in libraries, upstream TypeScript declarations and data-only URL/File helpers");
                using(var sample=new WorkbookDocument(sampleCatalog.Read("lodash"),_services))
                {
                    Check(!sample.ScriptsEnabled,"Opening a published sample leaves scripts disabled");
                    await Task.Run(()=>{sample.EnableScriptsAndCalculate();sample.RunAction("example");});
                    Check(sample.GetCellValue(2,1,2)=="200"&&sample.Undo()&&sample.GetCellValue(2,1,2)=="",
                        "Published sample executes its embedded TypeScript through the isolated worker and supports undo");
                }
                var action = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.ts", """
                    function main(workbook: QuickExcel.Workbook) {
                      if ('files' in host) throw Error('Unexpected host file API');
                      workbook.files.writeText('//xlsq/install-check.txt', 'verified');
                      const sheet=workbook.getActiveWorksheet(), range=sheet.getRange('B1:C3');
                      range.setValues([['Name','Value'],['b',2],['a',1]]);
                      range.getFormat().getFont().setBold(true);
                      range.getFormat().setHorizontalAlignment(QuickExcel.HorizontalAlignment.center);
                      const table=sheet.addTable(range,true); table.setName('InstallTable');
                      table.getSort().apply([{key:1}]);
                      table.getAutoFilter().apply(table.getRange(),1,{filterOn:QuickExcel.FilterOn.values,values:['1']});
                      if(!table.getAutoFilter().getIsDataFiltered()||table.getRangeBetweenHeaderAndTotal().getValues()[0][1]!==1)throw Error('P1 verification');
                      const chart=sheet.addChart(QuickExcel.ChartType.columnClustered,range);chart.setName('InstallChart');chart.getTitle().setText('Installation');
                      const conditional=sheet.getRange('C2:C3').addConditionalFormat(QuickExcel.ConditionalFormatType.custom);
                      conditional.getCustom()!.getRule().setFormula('=C2>0');conditional.getCustom()!.getFormat().getFont().setBold(true);
                      sheet.getRange('E1:E2').getDataValidation().setRule({list:{source:'yes,no',inCellDropDown:true}});
                      const comment=sheet.addComment('F1','Installation check');comment.addCommentReply('Verified');comment.setResolved(true);
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
                var objectSheet = Document.Snapshot.Sheets[0];
                Check(objectSheet.Data.Charts.Length == 1 && objectSheet.Data.ConditionalFormats.Length == 1 && objectSheet.Data.Validations.Length == 1 &&
                    objectSheet.Data.ThreadComments.Length == 2, "Published P2 chart, conditional format, validation and threaded comment transaction");
                var objectFile = Path.Combine(directory, "p2-objects.xlsq"); var objectSnapshot = Document.Snapshot;
                var reopened = await Task.Run(() => { var storage = new TypedWorkbookStorage(); storage.Save(objectFile, objectSnapshot); return storage.Open(objectFile); });
                Check(reopened.Sheets[0].Data.ThreadComments.Length == 2 && reopened.Sheets[0].Data.Charts[0].Name == "InstallChart",
                    "Published P2 object save/reopen");
                Check(reopened.PackageFiles.Entries.ContainsKey("//xlsq/install-check.txt") && Document.Undo() &&
                    !Document.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/install-check.txt"),
                    "Published API v3 workbook.files saves and reopens virtual files and undoes the action");
            }
            var visualSheet = Document.Snapshot.ActiveSheetId;
            if (OperatingSystem.IsWindows())
            {
                var folder = Directory.CreateDirectory(Path.Combine(directory, "script-files")).FullName;
                var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                {
                    ["main"] = new("scripts/main.js", "export function main() {}"),
                    ["files"] = new("scripts/files.js", """
                        export async function main(w,h) {
                          const dir = await h.fs.getUserSelectedDirectory({fileAccess:'rw'});
                          const file = await dir.getFile('runtime.txt',{fileMode:'create'});
                          const output = await file.open();
                          output.writeLine('QuickExcel 文件');
                          await output.close();
                          let found = false;
                          for await (const child of dir.entries()) if (child.name === file.name && child.type === 'file') found = true;
                          if (!found) throw Error('Selected directory entry missing');
                          const input = await file.open({buffering:'stream'});
                          w.getWorksheets()[0].getRange('A1').setValue(await input.readLine());
                          await input.close();
                          for (const extension of ['xlsx','xlsq']) {
                            const [target] = await h.fs.getUserSelectedFiles({mode:'save', fileMode:'create',
                              fileAccess:'rw', suggestedName:'selected.' + extension});
                            const book = await target.openWorkbook();
                            const sheet = book.workbook.getWorksheet('Sheet1');
                            sheet.getRange('A1:B1').setValues([[6,7]]);
                            sheet.getRange('C1').setFormula('=A1*B1');
                            sheet.getRange('C1').getFormat().getFont().setBold(true);
                            await book.close();
                            const saved = await target.openWorkbook();
                            const result = saved.workbook.getWorksheet('Sheet1').getRange('C1');
                            if (result.getValue() !== 42 || !result.getFormat().getFont().getBold())
                              throw Error('Selected workbook formula or format did not round trip');
                            if (extension === 'xlsx') {
                              const copy = await dir.getFile('report-copy.xlsq',{fileMode:'create'});
                              await saved.saveAs(copy,'xlsq');
                            }
                            await saved.close();
                          }
                          await w.saveAs(await dir.getFile('active-copy.xlsx',{fileMode:'create'}),'xlsx');
                        }
                        """)
                }, ["main", "files"], scriptApiVersion: 4, bindings: [new("main", "initialization"), new("files", "actions")]);
                using var files = new WorkbookDocument(Document.Snapshot.WithScriptModules(modules), _services);
                files.Session.UserFilePicker = (request, begin, _) =>
                {
                    using var interaction = begin();
                    var selected = request.Directory ? folder : request.SuggestedName switch
                    {
                        "selected.xlsx" or "selected.xlsq" => Path.Combine(folder, request.SuggestedName),
                        _ => throw new InvalidOperationException("Unexpected self-check file selection.")
                    };
                    return Task.FromResult(new ScriptFileSelection([selected], file));
                };
                await Task.Run(() => { files.EnableScriptsAndCalculate(); files.RunAction("files"); });
                Check(File.ReadAllText(Path.Combine(folder, "runtime.txt")) == "QuickExcel 文件\n" &&
                    files.GetCellValue(0, 0, 0) == "QuickExcel 文件",
                    "Published v4 host.fs writes a selected file and streams it back through the isolated worker");
                Check(new[] { "selected.xlsx", "selected.xlsq" }.All(name =>
                    _services.Storage.Open(Path.Combine(folder, name)).Sheets[0].GetCell(0, 2)?.Content.FormulaText == "=A1*B1"),
                    "Published v4 FileHandle.openWorkbook writes and reopens selected XLSX and XLSQ formulas");
                Check(_services.Storage.Open(Path.Combine(folder, "report-copy.xlsq")).Sheets[0].GetCell(0, 2)?.Content.FormulaText == "=A1*B1" &&
                    _services.Storage.Open(Path.Combine(folder, "active-copy.xlsx")).Sheets[0].GetCell(0, 0)?.Value.TextValue == "QuickExcel 文件",
                    "Published v4 saveAs exports a selected workbook and the live Action draft");
            }
            if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
            {
                var startup = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                {
                    ["main"] = new("scripts/main.js", "export function main(){ globalThis.count=40; }"),
                    ["action"] = new("scripts/action.js", "export function main(w,h){ if('functions' in h || 'events' in h) throw Error('manual registration'); w.getActiveWorksheet().getRange('G1').setValue(++globalThis.count); }"),
                    ["events"] = new("scripts/events.js", "export function onCellClick(c){ c.workbook.getActiveWorksheet().getRange('H1').setValue(++globalThis.count); }"),
                    ["formula"] = new("scripts/formula.ts", "export function DOUBLE(value: unknown): number { if(typeof value !== 'number') throw Error('number required'); return value*2; }", ScriptLanguage.TypeScript)
                }, ["main", "action", "events", "formula"], scriptApiVersion: 4,
                    bindings: [new("main", "initialization"), new("action", "actions"), new("events", "events"), new("formula", "formula-functions")]);
                using var initialized = new WorkbookDocument(Document.Snapshot.WithScriptModules(startup), _services);
                initialized.SetCell(0, 0, 5, "=JS.formula.DOUBLE(21)");
                await Task.Run(() => initialized.EnableScriptsAndCalculate());
                await Task.Run(() => initialized.RunAction("action"));
                initialized.Session.RaiseScriptEvent("cell.click", initialized.Snapshot.ActiveSheetId);
                await Task.Run(() => initialized.DispatchScriptEvents());
                await Task.Run(() => initialized.Calculate());
                await Task.Run(() => initialized.RunAction("action"));
                Check(initialized.GetCellValue(0, 0, 5) == "42" && initialized.GetCellValue(0, 0, 6) == "43" && initialized.GetCellValue(0, 0, 7) == "42",
                    "Published API v4 initializes once, discovers UDF/Actions/Events and preserves isolated state across recalculation");
                var startupPath = Path.Combine(directory, "initialized-scripts.xlsq"); initialized.Save(startupPath);
                Check(_services.Storage.Open(startupPath).ScriptModules.ScriptApiVersion == 4, "Published API v4 roles and sources save/reopen");
                await Task.Run(() => initialized.RunConsole("localStorage.set('installation','verified'); sessionStorage.set('installation','memory'); if(localStorage.entries()[0][1]!=='verified')throw Error('storage');"));
                Check(initialized.Snapshot.LocalStorage.Get("installation") == "verified" && initialized.Session.SessionStorage.Get("installation") == "memory",
                    "Published IStorage globals execute in the isolated transaction");
                Check(initialized.Undo() && initialized.Snapshot.LocalStorage.Get("installation") is null && initialized.Session.SessionStorage.Get("installation") == "memory" && initialized.Redo(),
                    "Published localStorage undo/redo preserves sessionStorage");
                initialized.Save(startupPath);
                using var stored = new WorkbookDocument(_services.Storage.Open(startupPath), _services);
                Check(stored.Snapshot.LocalStorage.Get("installation") == "verified" && stored.Session.SessionStorage.Entries.Count == 0,
                    "Published storage persists local data and releases session data on reopen");
            }
            _workspace.Editor.Execute(new SetCellsCommand(Enumerable.Range(0, 3).SelectMany(column => Enumerable.Range(10, 4).Select(row =>
                new CellEdit(new(visualSheet, row, column), CellContent.FromLiteral(CellValue.FromNumber((row - 11) * 10)))))));
            _workspace.Editor.Execute(new SetConditionalFormatsCommand(visualSheet, [
                ConditionalVisualPresets.DataBar(new(10, 0, 13, 0), "FF638EC6", false),
                ConditionalVisualPresets.ColorScale(new(10, 1, 13, 1), ["FFF8696B", "FFFFEB84", "FF63BE7B"]).WithPriority(2),
                ConditionalVisualPresets.IconSet(new(10, 2, 13, 2), "3Arrows").WithPriority(3)]));
            var visuals = Document.Snapshot; var visualModel = visuals.GetSheet(visualSheet);
            Check(ConditionalPresentation.GetVisual(visuals, visualModel, 10, 0)?.Bar is { Gradient: false, Color: "FFFF0000" } &&
                ConditionalPresentation.GetStyle(visuals, visualModel, 10, 1).FillColor == "FFF8696B" &&
                ConditionalPresentation.GetVisual(visuals, visualModel, 13, 2)?.Icon?.Index == 2, "Published data bars, color scales and icon sets project correctly");
            var visualPath = Path.Combine(directory, "visual-conditionals.xlsq"); var visualStorage = new TypedWorkbookStorage();
            visualStorage.Save(visualPath, visuals); var visualReopened = visualStorage.Open(visualPath);
            Check(visualModel.Data.ConditionalFormats.Zip(visualReopened.Sheets[0].Data.ConditionalFormats).All(p => p.First.ContentEquals(p.Second)) &&
                visualReopened.Sheets[0].Data.ConditionalFormats.Length == 3, "Published visual conditional rules save/reopen");
            Synchronize();
            for (var wait = 0; wait < 250 && _execution.IsRunning; wait++) await Task.Delay(20);
            Check(!_execution.IsRunning, "Pending calculation settled before print preview");
            await OpenPrintPreviewAsync(); Check(_pdfPage is not null && _printSession is not null, "Production print preview was created");
            Check(_pdfPage!.Browser.IsReady && _printSession!.Job.Layout.Pages.Length > 0, "Production PDF.js worker and complete print-session PDF");
            await _printSession!.ExportAsync(Path.Combine(directory, "installation.pdf"));
            await ClosePrintPreviewAsync(); Check(true, "PDF export and browser/resource cleanup");
            await _workspace.OpenAsync(new(file), readOnly: true);
            Check(Document.IsReadOnly && !Document.IsDirty && _grid.IsCellReadOnly(_grid.Current) && _readOnlyNotice.IsVisible,
                "Published read-only mode protects the grid and displays its state");
            var blocked = false;
            try { Document.SetCell(0, 0, 0, "must not change"); }
            catch (InvalidOperationException) { blocked = true; }
            Check(blocked && !Document.IsDirty && Document.GetCellInput(0, 0, 0) == "安装检查 中文 123", "Read-only session rejects edits without changing the source");
            for (var wait = 0; wait < 250 && _execution.IsRunning; wait++) await Task.Delay(20);
            Check(!_execution.IsRunning, "Read-only workbook finished its pending calculation before print preview");
            await OpenPrintPreviewAsync();
            Check(_pdfPage is not null && _printSession is not null, "Read-only workbook print preview was created");
            Check(_pdfPage!.Browser.IsReady && _printSession!.Job.Layout.Pages.Length > 0, "Read-only workbook retains print preview");
            await ClosePrintPreviewAsync();
            await _workspace.SaveAsync(new(Path.Combine(directory, "editable-copy.xlsq")));
            Check(!Document.IsReadOnly && !Document.IsDirty && !_grid.IsReadOnly, "Saving a writable copy restores editing");
            var recoveryStore = new WorkbookRecoveryStore(Path.Combine(directory, "recovery"));
            using (var lease = recoveryStore.Start(Document.Session.SessionId, Document.Snapshot.DocumentId))
            {
                await Task.Run(() => lease.Capture(Document.Snapshot, _workspace.FilePath, _workspace.Version, DateTimeOffset.UtcNow));
                Check(recoveryStore.Scan().Items.Count == 0, "Published recovery excludes the live session lease");
            }
            var recoveryItem = recoveryStore.Scan().Items.Single(); var recovered = recoveryStore.Recover(recoveryItem);
            using (var recoveredDocument = new WorkbookDocument(recovered, _services))
                Check(!recoveredDocument.ScriptsEnabled && recoveredDocument.GetCellInput(0, 0, 0) == "安装检查 中文 123", "Published recovery restores typed content without script authorization");
            recoveryStore.Discard(recoveryItem); Check(recoveryStore.Scan().Items.Count == 0, "Published recovery discard removes completed generations");
        }
        catch (Exception ex) { failure = ex; }
        finally { try { await CloseScriptWindowAsync(); await ClosePrintPreviewAsync(); } catch (Exception ex) { failure ??= ex; } }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new SelfCheckResult(failure is null, checks, failure?.ToString(), WebRuntimePrerequisite.Version,
            Program.SoftwareRendering, Environment.OSVersion.ToString(), System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription), DesktopJsonContext.Indented.SelfCheckResult));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
