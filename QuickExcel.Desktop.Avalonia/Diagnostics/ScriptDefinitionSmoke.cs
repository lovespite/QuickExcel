using System.Text.Json;
using System.Text.RegularExpressions;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckColdJavaScriptDefinitionsAsync(Func<object, Task<JsonElement>> test, Action<bool, string> check)
    {
        var originals = Document.Snapshot.ScriptModules;
        foreach (var name in new[] { "quickexcel.d.ts", "lib.es5.d.ts" })
        {
            // Each case starts with a fresh page containing only JavaScript models.
            // Opening a declaration must not require an earlier TypeScript module.
            if (!await NavigateScriptAsync(new(null, name)))
                throw new InvalidOperationException($"Cold JavaScript definition navigation failed ({name}): {_scriptStatus!.Text}");
            var state = await test(new { action = "assistance" });
            check(!_scriptPage!.Drafts.Failed && state.GetProperty("definition").GetString() == name &&
                state.GetProperty("readOnly").GetBoolean() && state.GetProperty("source").GetString()!
                    .Contains(name == "quickexcel.d.ts" ? "declare const host" : "interface Array<T>"),
                $"A fresh JavaScript-only page opens {name} read-only without requiring TypeScript registration");
            check(await NavigateScriptAsync(new("main", null)) && await CommitScriptsAsync() &&
                ReferenceEquals(originals, Document.Snapshot.ScriptModules),
                "Cold definition browsing returns to the original JavaScript source without changing the workbook");
            state = await test(new { action = "inspect" });
            check(state.GetProperty("language").GetString() == "javascript" && !state.GetProperty("readOnly").GetBoolean() &&
                state.GetProperty("source").GetString() == originals.Modules["main"].Source,
                "The JavaScript editor and language worker remain usable after cold definition navigation");
            await CloseScriptWindowAsync();
            await OpenScriptsAsync();
        }
    }

    private async Task CheckScriptDefinitionsAsync(Func<object, Task<JsonElement>> test, Action<bool, string> check, string directory)
    {
        var originals = Document.Snapshot.ScriptModules;
        async Task<JsonElement> Ui(object? payload = null) => await test(payload ?? new { action = "assistance" });
        async Task<JsonElement> Until(Func<JsonElement, bool> predicate, string description)
        {
            JsonElement state = default;
            for (var retry = 0; retry < 100; retry++)
            {
                state = await Ui();
                if (!_working && predicate(state)) { check(true, description); return state; }
                await Task.Delay(50);
            }
            await File.WriteAllTextAsync(Path.Combine(directory, "definition-failure.json"), state.ToString());
            throw new InvalidOperationException(description);
        }
        async Task Main()
        {
            _modulePicker!.SelectedItem = "main";
            await Until(s => s.GetProperty("definition").ValueKind == JsonValueKind.Null && s.GetProperty("selected").GetString() == "main",
                "The native file list returns from declarations to the editable workbook module");
        }
        async Task Edit(string text)
        {
            await Ui(new { action = "assistance", command = "editor.action.hideHover" });
            await test(new { action = "edit", module = "main", text });
        }

        _definitionPicker!.SelectedItem = "quickexcel.d.ts";
        var state = await Until(s => s.GetProperty("definition").GetString() == "quickexcel.d.ts" && s.GetProperty("readOnly").GetBoolean(),
            "The native definitions list opens the actual bundled QuickExcel declarations read-only");
        var declarations = state.GetProperty("source").GetString()!;
        check(declarations.Contains("declare const host") && !Regex.IsMatch(declarations, @"[\u4e00-\u9fff]"),
            "The definition viewer exposes the complete API with English documentation");
        await Ui(new { action = "assistance", typeText = "THIS MUST NOT BE INSERTED" });
        state = await Ui();
        check(state.GetProperty("source").GetString() == declarations && state.GetProperty("suggestDisabled").GetBoolean() &&
            await CommitScriptsAsync() && ReferenceEquals(originals, Document.Snapshot.ScriptModules) && _scriptPage!.Drafts.Models.Count == originals.Modules.Count,
            "Typing and draft flushing never modify, save or execute a definition model");
        await Main();

        const string invalid = "const invalid: QuickExcel.FunctionDescriptor = { name: 'JS.TEST' };";
        await Edit(invalid);
        await Until(s => s.GetProperty("markers").EnumerateArray().Any(m => m.GetProperty("code").ToString() == "2741"),
            "A missing required API descriptor property produces a diagnostic with declaration information");
        await Ui(new { action = "assistance", offset = 8, button = "editor.action.showHover" });
        state = await Until(s => s.GetProperty("hover").GetString()!.Contains("quickexcel.d.ts("),
            "The actual diagnostic hover contains a clickable quickexcel.d.ts related-information link");
        var location = Regex.Match(state.GetProperty("hover").GetString()!, @"quickexcel\.d\.ts\((\d+),\s*(\d+)\)");
        await Ui(new { action = "assistance", link = "quickexcel.d.ts(" });
        state = await Until(s => s.GetProperty("definition").GetString() == "quickexcel.d.ts" && _selectedDefinition == "quickexcel.d.ts",
            "Clicking the diagnostic link opens the definition and synchronizes the native navigation");
        check(state.GetProperty("position").GetProperty("lineNumber").GetInt32() == int.Parse(location.Groups[1].Value) &&
            state.GetProperty("position").GetProperty("column").GetInt32() == int.Parse(location.Groups[2].Value) &&
            Document.Snapshot.ScriptModules.Modules["main"].Source == invalid,
            "Diagnostic links preserve the pending source and land on the declared line and column");
        await Main();
        check(!(await Ui()).GetProperty("readOnly").GetBoolean(), "Returning from a definition restores source editability");
        const string api = "export function main(workbook: QuickExcel.Workbook) { workbook.getActiveCell(); }";
        await Edit(api);
        await Ui(new { action = "assistance", offset = api.IndexOf("getActiveCell", StringComparison.Ordinal) + 2, button = "editor.action.revealDefinition" });
        await Until(s => s.GetProperty("definition").GetString() == "quickexcel.d.ts" && s.GetProperty("position").GetProperty("lineNumber").GetInt32() > 1,
            "Go to Definition resolves QuickExcel API members through the actual Monaco provider");
        await Main();
        const string standard = "const values = [1, 2]; values.map(value => value * 2);";
        await Edit(standard);
        await Ui(new { action = "assistance", offset = standard.IndexOf(".map", StringComparison.Ordinal) + 2, button = "editor.action.revealDefinition" });
        await Until(s => s.GetProperty("definition").GetString() == "lib.es5.d.ts" && s.GetProperty("readOnly").GetBoolean() &&
            _scriptDefinitionFiles.Contains("lib.es5.d.ts"), "Standard ECMAScript declarations open locally, read-only, and remain available in the file list");
        await Main();
        await Edit(originals.Modules["main"].Source);
        await Ui(new { action = "assistance", offset = originals.Modules["main"].Source.LastIndexOf("twice", StringComparison.Ordinal) + 2, button = "editor.action.revealDefinition" });
        await Until(s => s.GetProperty("selected").GetString() == "helpers" && s.GetProperty("definition").ValueKind == JsonValueKind.Null &&
            _selectedModule == "helpers" && !s.GetProperty("readOnly").GetBoolean(),
            "Cross-module definition navigation selects the existing editable helper model and updates its native group");
        await Main();
        check(await CommitScriptsAsync() && Document.Snapshot.ScriptModules.Modules.Count == originals.Modules.Count &&
            Document.Snapshot.ScriptModules.Modules.All(p => p.Value.Source == originals.Modules[p.Key].Source),
            "Definition browsing adds no package modules and preserves every restored workbook source");
    }
}
