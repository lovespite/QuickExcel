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
        check(_modulePicker!.Items.Cast<string>().SequenceEqual(["main"]) && _scriptModuleGroups.All(g => g.Picker.Items.Count == 0),
            "The default main.ts stays above all three initially empty script categories");
        var initial = await test(new { action = "inspect" });
        check(initial.GetProperty("language").GetString() == "typescript" &&
            Document.Snapshot.ScriptModules.Modules["main"].Path == "scripts/main.ts" &&
            ScriptModuleProperties(Document.Snapshot.ScriptModules, "main") == "主入口 · TypeScript" &&
            _scriptStatus!.Text!.Contains("主入口 · TypeScript"),
            "New workbooks open a TypeScript main entry and show its entry role in the status and file properties");
        check(!_scriptWindow!.GetVisualDescendants().OfType<Button>().Any(b => b.Content is "模块角色" or "JS / TS"),
            "The workspace exposes no post-creation role or language editor");
        var beforeCancel = Document.Snapshot.ScriptModules;
        _dialogDriver = async window =>
        {
            var dialog = (FieldsDialog)window;
            check(dialog.String("id") == "" && dialog.Index("role") == 1 &&
                ((ComboBox)dialog.Input("role")).Items.Cast<string>().SequenceEqual(["Common", "Actions", "UDF"]),
                "Module creation starts with an empty ID and Actions selected from exactly Common, Actions and UDF");
            check(dialog.Index("language") == (int)ScriptLanguage.TypeScript,
                "The module creation dialog defaults to TypeScript");
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            using (var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(window.Bounds.Width), (int)Math.Ceiling(window.Bounds.Height)), new Vector(96, 96)))
            { image.Render(window); image.Save(Path.Combine(directory, "script-module-dialog.png"), PngBitmapEncoderOptions.Default); }
            window.Close(false);
        };
        await AddScriptModuleAsync();
        check(ReferenceEquals(beforeCancel, Document.Snapshot.ScriptModules), "Cancelling module creation preserves the complete module set");
        foreach (var (id, role, language) in new[] { ("utility", 0, 0), ("formula", 2, 1), ("action", 1, 0) })
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
            var inspected = await test(new { action = "inspect" });
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
        check(modules.EntryModules.SequenceEqual(["main", "formula", "action"]) && modules.Bindings.ToHashSet().SetEquals(
            [new("main", "formula-functions"), new("main", "actions"), new("formula", "formula-functions"), new("action", "actions")]),
            "Creation persists exclusive UDF or Actions bindings without turning Common modules into execution entries");
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
        await test(new { action = "edit", module = "main", text = "// draft before category navigation\n" });
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
