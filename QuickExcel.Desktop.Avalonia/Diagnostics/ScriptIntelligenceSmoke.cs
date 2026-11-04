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

        const string member = "host.functions";
        await Edit("main", member);
        await Ui(new { action = "assistance", offset = member.Length, typeText = "." });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("register"),
            "Typing a dot automatically opens the actual host completion list");
        state = await Ui(new { action = "assistance", command = "acceptSelectedSuggestion" });
        check(state.GetProperty("source").GetString() == member + ".register", "Accepting a suggestion inserts the selected API member into the draft");

        await Edit("main", "host");
        await Ui(new { action = "assistance", offset = 4, typeText = "." });
        await Until(s => new[] { "functions", "events", "files" }.All(name => s.GetProperty("suggestions").GetString()!.Contains(name)),
            "The host extension completes functions, events and files");
        await Edit("main", "host.functions.");
        await Ui(new { action = "assistance", offset = "host.functions.".Length, button = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("register"), "Host function extension exposes register completion");
        // The generated setValue signature accepts any; setValues constrains each cell's type.
        await Edit("main", "export function main(workbook: QuickExcel.Workbook) {\n  workbook.getActiveWorksheet().getRange('A1').setValues([[{}]]);\n}");
        await Until(s => HasMarker(s, "2322"), "Typed Range.setValues rejects object cells in live diagnostics");
        await Edit("main", "export function main(workbook: QuickExcel.Workbook) {\n  workbook.getActiveWorksheet().getRange('A1:C1').setValues([[42, 'text', true]]);\n}");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0,
            "Correcting range values to numbers, strings and booleans clears the diagnostic");

        const string parameters = "host.files.writeText('//temp/note.txt', 'hello');";
        await Edit("main", parameters);
        await Ui(new { action = "assistance", offset = parameters.IndexOf("writeText(", StringComparison.Ordinal) + "writeText(".Length, button = "editor.action.triggerParameterHints" });
        await Until(s => s.GetProperty("parameters").GetString()!.Contains("path") && s.GetProperty("parameters").GetString()!.Contains("text"),
            "Parameter-hint control renders the host file signature");

        const string hover = "host.files.writeText('//temp/note.txt', 'hello');";
        await Edit("main", hover);
        await Ui(new { action = "assistance", offset = hover.IndexOf("writeText", StringComparison.Ordinal) + 3, button = "editor.action.showHover" });
        await Until(s => s.GetProperty("hover").GetString()!.Contains("writeText") && s.GetProperty("hover").GetString()!.Contains("string"),
            "Hover control renders declared host member signatures");

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
        await Edit("main", "host.events.");
        await Ui(new { action = "assistance", offset = "host.events.".Length, button = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("on"), "Event registration appears in editor completion");
        await Edit("main", "host.events.on('cell.click', context => {\n  context.workbook.getActiveWorksheet().getRange('A1').setValues([[{}]]);\n  context.host.files.writeText('//temp/note.txt', 42);\n});");
        await Until(s => HasMarker(s, "2322") && HasMarker(s, "2345"), "Event context infers workbook cell types and host file argument types");
        await Edit("main", "host.events.on('cell.click', context => {\n  context.workbook.getActiveWorksheet().getRange('A1').setValues([[42]]);\n  context.host.files.writeText('//temp/note.txt', 'hello');\n});");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0,
            "Correcting event workbook and host arguments clears both diagnostics");
        foreach (var (id, module) in originals.Modules) await Edit(id, module.Source);
        await _scriptPage.SelectAsync("main");
    }
}
