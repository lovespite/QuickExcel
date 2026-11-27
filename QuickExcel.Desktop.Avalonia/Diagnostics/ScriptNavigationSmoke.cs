using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckScriptNavigationAsync(Func<object, Task<JsonElement>> test, Func<string, Task> select,
        Action<bool, string> check, string directory)
    {
        _workspace.New(); await OpenScriptsAsync();
        if (DesktopServices.SupportsScripts)
        {
            check(_scriptApplyButton is { IsVisible: true, IsEnabled: true, Content: global::Avalonia.Controls.Shapes.Path }, "The script editor exposes its Run and Apply button");
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            using (var image = new RenderTargetBitmap(new PixelSize(136, 120), new Vector(384, 384)))
            { image.Render(_scriptApplyButton!); image.Save(Path.Combine(directory, "script-run-button.png"), PngBitmapEncoderOptions.Default); }
            _scriptApplyButton!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _operation;
            check(Document.ScriptsEnabled, "The Run button applies scripts through the isolated execution workflow");
            DisableScripts();
            foreach (var request in new[] { new Dictionary<string, object?> { ["action"] = "applyShortcut", ["repeat"] = true, ["composing"] = false, ["control"] = false },
                new Dictionary<string, object?> { ["action"] = "applyShortcut", ["repeat"] = false, ["composing"] = true, ["control"] = false },
                new Dictionary<string, object?> { ["action"] = "applyShortcut", ["repeat"] = false, ["composing"] = false, ["control"] = true } })
            { await test(request); await _operation; check(!Document.ScriptsEnabled, "Modified, repeated or composing F5 does not apply scripts"); }
            var pressed = await test(new Dictionary<string, object?> { ["action"] = "applyShortcut" }); await _operation;
            check(pressed.GetProperty("prevented").GetBoolean() && Document.ScriptsEnabled, "F5 in Monaco applies scripts without refreshing the WebView");
            DisableScripts();
        }
        check(_scriptWindow!.GetVisualDescendants().OfType<Menu>().Any(menu => menu.Items.OfType<MenuItem>().Select(item => item.Header?.ToString()).SequenceEqual(["文件", "编辑", "视图", "运行", "签名"])),
            "The script workspace groups commands into File, Edit, View, Run and Signature menus");
        check(_modulePicker!.Items.Cast<string>().SequenceEqual(["main"]) && _scriptModuleGroups.All(g => g.Picker.Items.Count == 0),
            "The default main.ts stays above all four initially empty script categories");
        var initial = await test(new Dictionary<string, object?> { ["action"] = "inspect" });
        check(initial.GetProperty("language").GetString() == "typescript" &&
            Document.Snapshot.ScriptModules.Modules["main"].Path == "scripts/main.ts" &&
            ScriptModuleProperties(Document.Snapshot.ScriptModules, "main") == "主入口 · TypeScript" &&
            _scriptStatus!.Text!.Contains("主入口 · TypeScript"),
            "New workbooks open a TypeScript main entry and show its entry role in the status and file properties");
        var originalMain = initial.GetProperty("source").GetString()!;
        check(await NavigateScriptAsync(new(null, "quickexcel.d.ts")), "The v4 workbook opens its QuickExcel definition file");
        var apiDefinition = await test(new Dictionary<string, object?> { ["action"] = "assistance" });
        var declarations = apiDefinition.GetProperty("source").GetString()!;
        check(apiDefinition.GetProperty("readOnly").GetBoolean() &&
            declarations.Contains("interface Host extends V4HostExtensions") && declarations.Contains("getUserSelectedFiles(") &&
            declarations.Contains("openWorkbook(): Promise<SelectedWorkbook>") && declarations.Contains("interface SelectedWorkbook extends WorkbookSaveAs") &&
            declarations.Contains("readonly workbook: Selected.Workbook") && declarations.Contains("namespace Selected") &&
            declarations.Contains("interface Workbook extends WorkbookSaveAs") && declarations.Contains("saveAs(file: FileHandle"),
            "The v4 definition viewer includes selected.workbook and the file lifecycle methods");
        await NavigateScriptAsync(new("main", null));
        foreach (var (id,source) in new (string,string)[] {
            ("lodash", "import {sumBy} from 'lodash'; const n:number=sumBy([{n:1}],'n');"),
            ("decimal.js", "import Decimal from 'decimal.js'; const n:string=new Decimal('0.1').plus('0.2').toString();"),
            ("dayjs", "import dayjs from 'dayjs'; const n:number=dayjs.duration({minutes:1}).asMinutes(); dayjs.tz('2024-01-01','Asia/Shanghai');"),
            ("papaparse", "import Papa from 'papaparse'; const rows=Papa.parse<{n:number}>('n\\n1',{header:true,dynamicTyping:true}).data; const n:number=rows[0].n;"),
            ("simple-statistics", "import {mean} from 'simple-statistics'; const n:number=mean([1,2]);"),
            ("zod", "import {z} from 'zod'; const n:number=z.number().parse(1); const f:File=z.file().parse(new File(['x'],'x'));"),
            ("fuse.js", "import Fuse from 'fuse.js'; const n:number=new Fuse([{n:1,name:'a'}],{keys:['name']}).search('a')[0].item.n;"),
            ("arquero", "import {from,op} from 'arquero'; const n:number=from([{n:1}]).rollup({n:op.sum('n')}).get('n',0) as number;"),
            ("mathjs", "import {add} from 'mathjs'; const n:number=add(1,2);") })
        {
            await test(new Dictionary<string,object?> { ["action"]="edit",["module"]="main",["text"]=source });
            var typed=await test(new Dictionary<string,object?> { ["action"]="inspect" });
            await File.WriteAllTextAsync(Path.Combine(directory,"library-"+id+".json"),typed.ToString());
            check(typed.GetProperty("diagnostics").GetArrayLength()==0,"Monaco checks the bundled declarations for "+id);
        }
        var lodashDefinition=QuickExcel.ScriptContracts.BuiltInLibraryIdentity.EntryDefinitions["lodash"];
        check(await NavigateScriptAsync(new(null,lodashDefinition)),"Bundled library definitions are navigable through the host allowlist");
        var definition=await test(new Dictionary<string,object?> { ["action"]="inspect" });
        check(definition.GetProperty("readOnly").GetBoolean() && definition.GetProperty("source").GetString()!.Contains("LoDashStatic"),"Library definitions open read-only with upstream types");
        await NavigateScriptAsync(new("main",null));
        await test(new Dictionary<string,object?> { ["action"]="edit",["module"]="main",["text"]=originalMain });
        check(await CommitScriptsAsync(),"Library editor probes restore and commit the original draft");
        check(!_scriptWindow!.GetVisualDescendants().OfType<Button>().Any(b => b.Content is "模块角色" or "JS / TS"),
            "The workspace exposes no post-creation role or language editor");
        var beforeCancel = Document.Snapshot.ScriptModules;
        _dialogDriver = async window =>
        {
            var dialog = (FieldsDialog)window;
            check(dialog.String("id") == "" && dialog.Index("role") == 1 &&
                ((ComboBox)dialog.Input("role")).Items.Cast<string>().SequenceEqual(["Common", "Actions", "UDF", "Events"]),
                "Module creation starts with an empty ID and Actions selected from Common, Actions, UDF and Events");
            check(dialog.Index("language") == (int)ScriptLanguage.TypeScript,
                "The module creation dialog defaults to TypeScript");
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            using (var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(window.Bounds.Width), (int)Math.Ceiling(window.Bounds.Height)), new Vector(96, 96)))
            { image.Render(window); image.Save(Path.Combine(directory, "script-module-dialog.png"), PngBitmapEncoderOptions.Default); }
            window.Close(false);
        };
        await AddScriptModuleAsync();
        check(ReferenceEquals(beforeCancel, Document.Snapshot.ScriptModules), "Cancelling module creation preserves the complete module set");
        foreach (var (id, role, language) in new[] { ("utility", 0, 0), ("formula", 2, 1), ("action", 1, 0), ("events", 3, 1) })
        {
            _dialogDriver = window =>
            {
                var dialog = (FieldsDialog)window;
                ((TextBox)dialog.Input("id")).Text = id;
                if (role != 1) ((ComboBox)dialog.Input("role")).SelectedIndex = role;
                if (language == 0) ((ComboBox)dialog.Input("language")).SelectedIndex = language;
                window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, "确定")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return Task.CompletedTask;
            };
            await AddScriptModuleAsync();
            var created = Document.Snapshot.ScriptModules.Modules[id];
            var inspected = await test(new Dictionary<string, object?> { ["action"] = "inspect" });
            check(created.Language == (ScriptLanguage)language && created.Path == $"scripts/{id}.{(language == 0 ? "js" : "ts")}" &&
                inspected.GetProperty("selected").GetString() == id && inspected.GetProperty("language").GetString() == (language == 0 ? "javascript" : "typescript"),
                $"Creation fixes {id}'s role/language and opens the corresponding actual Monaco model");
        }
        var udf = _scriptModuleGroups.Single(g => g.Kind == "formula-functions");
        var actions = _scriptModuleGroups.Single(g => g.Kind == "actions");
        var common = _scriptModuleGroups.Single(g => g.Kind is null);
        check(udf.Picker.Items.Cast<string>().SequenceEqual(["formula"]) &&
            actions.Picker.Items.Cast<string>().SequenceEqual(["action"]) && common.Picker.Items.Cast<string>().SequenceEqual(["utility"]) &&
            common.Label == "Common" && ScriptModuleProperties(Document.Snapshot.ScriptModules, "utility") == "Common · JavaScript",
            "Created modules appear only in their selected category; Common has consistent labels and main stays ungrouped");
        var modules = Document.Snapshot.ScriptModules;
        await CheckFileSystemCompletionsAsync(test, check, modules.Modules["action"].Source, originalMain);
        check(modules.EntryModules.SequenceEqual(["main", "formula", "action", "events"]) && modules.Bindings.ToHashSet().SetEquals(
            [new("main", "initialization"), new("formula", "formula-functions"), new("action", "actions"), new("events", "events")]),
            "Creation persists exclusive UDF or Actions bindings without turning Common modules into execution entries");
        await CheckEventTemplateInsertionAsync(test, select, check, directory);
        await select("main");
        var revision = Document.Snapshot.Revision;
        foreach (var group in _scriptModuleGroups) group.Panel.IsExpanded = false;
        RefreshScriptChrome();
        await Task.Delay(250); // Capture the settled expander chevrons after their transition.
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        check(_scriptModuleGroups.All(g => !g.Panel.IsExpanded) && _modulePicker.IsVisible && _modulePicker.Bounds.Height > 0 &&
            Document.Snapshot.Revision == revision, "Collapsing categories keeps main visible and does not edit the workbook or reset on refresh");
        var navigation = (Control)_modulePicker.Parent!;
        using (var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(navigation.Bounds.Width), (int)Math.Ceiling(navigation.Bounds.Height)), new Vector(96, 96)))
        { image.Render(navigation); image.Save(Path.Combine(directory, "script-navigation-collapsed.png"), PngBitmapEncoderOptions.Default); }
        common.Panel.IsExpanded = true;
        await test(new Dictionary<string, object?> { ["action"] = "edit", ["module"] = "main", ["text"] = "// draft before category navigation\n" });
        await select("utility");
        check(Document.Snapshot.ScriptModules.Modules["main"].Source.Contains("draft before category navigation") &&
            _modulePicker.SelectedItem is null && common.Picker.SelectedItem is "utility" && !udf.Panel.IsExpanded && !actions.Panel.IsExpanded,
            "Cross-category selection flushes drafts, clears the prior highlight and retains unrelated collapsed panels");
        _scriptDiagnostics!.ItemsSource = new[] { new ScriptDiagnostic("formula location", Document.Snapshot.ScriptModules.Revision, "formula") };
        _scriptDiagnostics.SelectedIndex = 0;
        await ActivateScriptDiagnosticAsync();
        check(_selectedModule == "formula" && udf.Panel.IsExpanded && udf.Picker.SelectedItem is "formula" && common.Picker.SelectedItem is null,
            "Diagnostic navigation reveals the target in a collapsed category and synchronizes selection");
        udf.Panel.IsExpanded = false;
        await ActivateScriptDiagnosticAsync();
        check(udf.Panel.IsExpanded && _selectedModule == "formula", "Repeated diagnostic navigation also reveals an already selected but collapsed module");
        var beforeInvalid = Document.Snapshot.ScriptModules;
        _dialogDriver = window =>
        {
            var dialog = (FieldsDialog)window;
            ((ComboBox)dialog.Input("language")).SelectedIndex = (int)ScriptLanguage.JavaScript;
            foreach (var id in new[] { "", "   ", "main", "../invalid", "UTILITY" })
            {
                ((TextBox)dialog.Input("id")).Text = id;
                window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, "确定")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                check(window.IsVisible && ReferenceEquals(beforeInvalid, Document.Snapshot.ScriptModules),
                    $"Creation rejects empty, duplicate, invalid or path-colliding ID '{id}' in the dialog without mutating the workbook");
            }
            window.Close(false); return Task.CompletedTask;
        };
        await AddScriptModuleAsync();
        _files = new SmokeFiles { NextSave = Path.Combine(directory, "script-roles-roundtrip.xlsq") };
        check(await SaveAsync(false), "Created script roles and languages save through the regular workbook command");
        var saved = _services.Storage.Open(Path.Combine(directory, "script-roles-roundtrip.xlsq")).ScriptModules;
        check(saved.Bindings.ToHashSet().SetEquals(modules.Bindings) && saved.Modules.All(p => p.Value.Language == modules.Modules[p.Key].Language && p.Value.Path == modules.Modules[p.Key].Path),
            "Saving and reopening preserve every creation-time role, language and extension");
        foreach (var group in _scriptModuleGroups) group.Panel.IsExpanded = true;
        await Task.Delay(250);
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        using (var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(navigation.Bounds.Width), (int)Math.Ceiling(navigation.Bounds.Height)), new Vector(96, 96)))
        { image.Render(navigation); image.Save(Path.Combine(directory, "script-navigation-expanded.png"), PngBitmapEncoderOptions.Default); }
        await CloseScriptWindowAsync();
    }
}
