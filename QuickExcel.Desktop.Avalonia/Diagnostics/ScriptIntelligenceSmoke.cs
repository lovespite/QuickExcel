using System.Text.Json;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckScriptIntelligenceAsync(Func<object, Task<JsonElement>> test, Action<bool, string> check, string directory)
    {
        var originals = Document.Snapshot.ScriptModules;
        async Task<JsonElement> Ui(object? command = null) => await test(command ?? new { action = "assistance" });
        async Task<JsonElement> Until(Func<JsonElement, bool> predicate, string description)
        {
            for (var retry = 0; retry < 100; retry++)
            {
                var state = await Ui();
                if (predicate(state)) { check(true, description); return state; }
                await Task.Delay(50);
                if (retry == 99) await File.WriteAllTextAsync(Path.Combine(directory, "assistance-failure.json"), state.ToString());
            }
            throw new InvalidOperationException(description);
        }
        static bool HasMarker(JsonElement state, string code) => state.GetProperty("markers").EnumerateArray().Any(m => m.GetProperty("code").ToString() == code);
        async Task Edit(string module, string text)
        {
            foreach (var command in new[] { "hideSuggestWidget", "closeParameterHints", "editor.action.hideHover", "closeMarkersNavigation" })
                await Ui(new { action = "assistance", command });
            await test(new { action = "edit", module, text });
        }
        var state = await Ui();
        await File.WriteAllTextAsync(Path.Combine(directory, "assistance-actions.json"), state.ToString());
        var actions = state.GetProperty("actions").EnumerateArray().Select(v => v.GetString()).ToHashSet();
        check(new[] { "editor.action.triggerSuggest", "editor.action.triggerParameterHints", "editor.action.showHover", "editor.action.marker.next" }.All(actions.Contains),
            "Monaco installs completion, parameter, hover and diagnostic UI actions, not only worker providers");

        await Edit("main", "export const value: number = '类型错误';");
        state = await Until(s => HasMarker(s, "2322") && s.GetProperty("status").GetString()!.Contains("1 个错误"),
            "Live TypeScript semantic markers and the current-module error count update without applying or executing source");
        await Ui(new { action = "assistance", offset = 0, button = "editor.action.marker.next" });
        state = await Until(s => s.GetProperty("problems").GetString()!.Contains("2322"), "Next-problem control opens the actual diagnostic widget");
        check(state.GetProperty("position").GetProperty("column").GetInt32() > 1, "Diagnostic navigation moves the editor caret to the error");
        await Edit("main", "export const value: number = 42;");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0 && s.GetProperty("status").GetString()!.Contains("0 个错误"),
            "Correcting a type error removes stale markers and updates the status");
        await Edit("main", "export const value: = 42;");
        await Until(s => HasMarker(s, "1110"), "TypeScript syntax errors appear as live editor markers");

        const string member = "quickExcel.functions";
        await Edit("main", member);
        await Ui(new { action = "assistance", offset = member.Length, typeText = "." });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("register"),
            "Typing a dot automatically opens the actual QuickExcel completion list");
        state = await Ui(new { action = "assistance", command = "acceptSelectedSuggestion" });
        check(state.GetProperty("source").GetString() == member + ".register", "Accepting a suggestion inserts the selected API member into the draft");

        await Edit("main", "$");
        await Ui(new { action = "assistance", offset = 1, typeText = "." });
        await Until(s => new[] { "functions", "actions", "events" }.All(name => s.GetProperty("suggestions").GetString()!.Contains(name)),
            "Dollar alias automatically completes the same functions, actions and events API");
        await Edit("main", "$.functions.");
        await Ui(new { action = "assistance", offset = "$.functions.".Length, button = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("register"), "Dollar function alias exposes register completion");
        await Edit("main", "$.actions.register('Fill', context => { context.patch.setValue('sheet', 'A1', {}); });");
        await Until(s => HasMarker(s, "2345"), "Dollar action alias infers context types and rejects invalid patch values");

        const string parameters = "$.actions.register('Fill', context => {\n  context.patch.setValue('sheet', 'A1', 42);\n});";
        await Edit("main", parameters);
        await Ui(new { action = "assistance", offset = parameters.IndexOf("setValue(", StringComparison.Ordinal) + "setValue(".Length, button = "editor.action.triggerParameterHints" });
        await Until(s => s.GetProperty("parameters").GetString()!.Contains("sheetId") && s.GetProperty("parameters").GetString()!.Contains("QxScalar"),
            "Parameter-hint control renders the contextual QuickExcel patch signature");

        const string hover = "quickExcel.actions.register('Fill', context => {\n  context.fs.writeText('//temp/note.txt', 'hello');\n});";
        await Edit("main", hover);
        await Ui(new { action = "assistance", offset = hover.IndexOf("writeText", StringComparison.Ordinal) + 3, button = "editor.action.showHover" });
        await Until(s => s.GetProperty("hover").GetString()!.Contains("writeText") && s.GetProperty("hover").GetString()!.Contains("UTF-8"),
            "Hover control renders inferred API types and declaration documentation");

        await Edit("helpers", "export const tools = { twice: (n: number): number => n * 2 };");
        const string crossModule = "import {tools} from 'helpers';\ntools.";
        await Edit("main", crossModule);
        await Ui(new { action = "assistance", offset = crossModule.Length, button = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("twice"),
            "Completion resolves current drafts through declared logical module IDs");
        await Edit("main", "import {tools} from 'helpers';\nexport const result: number = tools.twice('bad');");
        await Until(s => HasMarker(s, "2345"), "Cross-module TypeScript argument types are checked live");
        await Edit("helpers", "export const tools = { twice: (n: string): number => n.length * 2 };");
        await _scriptPage!.SelectAsync("main");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0,
            "Changing a dependency draft refreshes importing-module diagnostics");
        await Edit("main", "document.querySelector('body');\nprocess.exit(0);\nexport const n: number = null;");
        await Until(s => HasMarker(s, "2322") && s.GetProperty("markers").EnumerateArray().Any(m => m.GetProperty("message").GetString()!.Contains("document")) &&
            s.GetProperty("markers").EnumerateArray().Any(m => m.GetProperty("message").GetString()!.Contains("process")),
            "Editor keeps strict null checks and excludes DOM and Node globals");
        foreach (var (id, module) in originals.Modules) await Edit(id, module.Source);
        await _scriptPage.SelectAsync("main");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0, "Restored TypeScript modules clear diagnostics before isolated compilation checks");
        check(!Document.ScriptsEnabled && ReferenceEquals(originals, Document.Snapshot.ScriptModules),
            "Language assistance neither executes scripts nor commits per-keystroke document changes");
        await Edit("main", "QxEventNames.");
        await Ui(new { action = "assistance", offset = "QxEventNames.".Length, button = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("ON_"), "Event enum has real editor completion");
        await Edit("main", "$.events.onWorksheet");
        await Ui(new { action = "assistance", offset = "$.events.onWorksheet".Length, button = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("onWorksheetActivate"), "Property event handlers appear in editor completion");
        await Edit("main", "$.events.onWorksheetActivate = (sheet, context) => { context.patch.setValue(sheet.id, 'A1', {}); };");
        await Until(s => HasMarker(s, "2345"), "Property handler infers worksheet and action context types");
        foreach (var (id, module) in originals.Modules) await Edit(id, module.Source);
        await _scriptPage.SelectAsync("main");
    }
}
