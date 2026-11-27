using System.Text.Json;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckScriptIntelligenceAsync(Func<object, Task<JsonElement>> test, Action<bool, string> check, string directory)
    {
        var originals = Document.Snapshot.ScriptModules;
        async Task<JsonElement> Ui(object? command = null) => await test(command ?? new Dictionary<string, object?> { ["action"] = "assistance" });
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
                await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["command"] = command });
            await test(new Dictionary<string, object?> { ["action"] = "edit", ["module"] = module, ["text"] = text });
        }
        var state = await Ui();
        await File.WriteAllTextAsync(Path.Combine(directory, "assistance-actions.json"), state.ToString());
        var actions = state.GetProperty("actions").EnumerateArray().Select(v => v.GetString()).ToHashSet();
        check(new[] { "editor.action.triggerSuggest", "editor.action.triggerParameterHints", "editor.action.showHover", "editor.action.marker.next" }.All(actions.Contains),
            "Monaco installs completion, parameter, hover and diagnostic UI actions, not only worker providers");

        await Edit("main", "export const value: number = '类型错误';");
        state = await Until(s => HasMarker(s, "2322") && s.GetProperty("status").GetString()!.Contains("1 个错误"),
            "Live TypeScript semantic markers and the current-module error count update without applying or executing source");
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = 0, ["button"] = "editor.action.marker.next" });
        state = await Until(s => s.GetProperty("problems").GetString()!.Contains("2322"), "Next-problem control opens the actual diagnostic widget");
        check(state.GetProperty("position").GetProperty("column").GetInt32() > 1, "Diagnostic navigation moves the editor caret to the error");
        await Edit("main", "export const value: number = 42;");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0 && s.GetProperty("status").GetString()!.Contains("0 个错误"),
            "Correcting a type error removes stale markers and updates the status");
        await Edit("main", "export const value: = 42;");
        await Until(s => HasMarker(s, "1110"), "TypeScript syntax errors appear as live editor markers");

        const string member = "host.functions";
        await Edit("main", member);
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = member.Length, ["typeText"] = "." });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("register"),
            "Typing a dot automatically opens the actual host completion list");
        state = await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["command"] = "acceptSelectedSuggestion" });
        check(state.GetProperty("source").GetString() == member + ".register", "Accepting a suggestion inserts the selected API member into the draft");

        await Edit("main", "host");
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = 4, ["typeText"] = "." });
        await Until(s => new[] { "functions", "events" }.All(name => s.GetProperty("suggestions").GetString()!.Contains(name)) &&
            !s.GetProperty("suggestions").GetString()!.Contains("files") && !s.GetProperty("suggestions").GetString()!.Contains("fs"),
            "API v3 host completion exposes events and functions and omits both file APIs");
        const string workbookMembers = "export function main(workbook: QuickExcel.Workbook) {\n  workbook.fi\n}";
        await Edit("main", workbookMembers);
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = workbookMembers.IndexOf("workbook.fi", StringComparison.Ordinal) + "workbook.fi".Length, ["button"] = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("files"), "Workbook member completion exposes virtual files");
        await Edit("main", "host.functions.");
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = "host.functions.".Length, ["button"] = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("register"), "Host function extension exposes register completion");
        // The generated setValue signature accepts any; setValues constrains each cell's type.
        await Edit("main", "export function main(workbook: QuickExcel.Workbook) {\n  workbook.getActiveWorksheet().getRange('A1').setValues([[{}]]);\n}");
        await Until(s => HasMarker(s, "2322"), "Typed Range.setValues rejects object cells in live diagnostics");
        await Edit("main", "export function main(workbook: QuickExcel.Workbook) {\n  workbook.getActiveWorksheet().getRange('A1:C1').setValues([[42, 'text', true]]);\n}");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0,
            "Correcting range values to numbers, strings and booleans clears the diagnostic");

        const string parameters = "export function main(workbook: QuickExcel.Workbook) {\n  workbook.files.writeText('//temp/note.txt', 'hello');\n}";
        await Edit("main", parameters);
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = parameters.IndexOf("writeText(", StringComparison.Ordinal) + "writeText(".Length, ["button"] = "editor.action.triggerParameterHints" });
        await Until(s => s.GetProperty("parameters").GetString()!.Contains("path") && s.GetProperty("parameters").GetString()!.Contains("text"),
            "Parameter-hint control renders the workbook file signature");

        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["command"] = "closeParameterHints" });
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = parameters.IndexOf("writeText", StringComparison.Ordinal) + 3, ["button"] = "editor.action.showHover" });
        await Until(s => s.GetProperty("hover").GetString()!.Contains("writeText") && s.GetProperty("hover").GetString()!.Contains("string"),
            "Hover control renders the declared workbook file member signature");

        await Edit("helpers", "export const tools = { twice: (n: number): number => n * 2 };");
        const string crossModule = "import {tools} from 'helpers';\ntools.";
        await Edit("main", crossModule);
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = crossModule.Length, ["button"] = "editor.action.triggerSuggest" });
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
        await Ui(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = "host.events.".Length, ["button"] = "editor.action.triggerSuggest" });
        await Until(s => s.GetProperty("suggestions").GetString()!.Contains("on"), "Event registration appears in editor completion");
        await Edit("main", "host.events.on('cell.click', context => {\n  context.workbook.getActiveWorksheet().getRange('A1').setValues([[{}]]);\n  context.workbook.files.writeText('//temp/note.txt', 42);\n});");
        await Until(s => HasMarker(s, "2322") && HasMarker(s, "2345"), "Event context infers workbook cell and file argument types");
        await Edit("main", "host.events.on('cell.click', context => {\n  context.workbook.getActiveWorksheet().getRange('A1').setValues([[42]]);\n  context.workbook.files.writeText('//temp/note.txt', 'hello');\n});");
        await Until(s => s.GetProperty("markers").GetArrayLength() == 0,
            "Correcting the event's cell and file arguments clears both diagnostics");
        foreach (var (id, module) in originals.Modules) await Edit(id, module.Source);
        await _scriptPage.SelectAsync("main");
    }

    private static async Task CheckFileSystemCompletionsAsync(Func<object, Task<JsonElement>> test, Action<bool, string> check,
        string actionSource, string mainSource)
    {
        Task<JsonElement> Edit(string module, string source) => test(new Dictionary<string, object?> {
            ["action"] = "edit", ["module"] = module, ["text"] = source });
        async Task CheckWidget(string widget, string description, params string[] expected)
        {
            for (var attempt = 0; attempt < 100; attempt++)
            {
                var state = await test(new Dictionary<string, object?> { ["action"] = "assistance" });
                if (expected.All(state.GetProperty(widget).GetString()!.Contains)) { check(true, description); return; }
                await Task.Delay(50);
            }
            throw new InvalidOperationException(description);
        }

        // Type into the generated JavaScript function. Its local host argument needs its own type information.
        var js = actionSource.ReplaceLineEndings("\n");
        var insertion = js.LastIndexOf('}');
        await Edit("action", js.Insert(insertion, "    host.fs\n"));
        await test(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = insertion + "    host.fs".Length, ["typeText"] = "." });
        await CheckWidget("suggestions", "Typing host.fs. in a new JavaScript Action shows both file and directory pickers",
            "getUserSelectedFiles", "getUserSelectedDirectory");
        await test(new Dictionary<string, object?> { ["action"] = "assistance", ["command"] = "hideSuggestWidget" });
        var asyncJs = js.Replace("export function main(", "export async function main(", StringComparison.Ordinal);
        var picker = "    const [file] = await host.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'rw'});\n";
        foreach (var (before, expected) in new[] {
            (picker + "    file", new[] { "openWorkbook" }),
            (picker + "    const book = await file.openWorkbook();\n    book", new[] { "getValues", "saveAs", "setFormulas" }) })
        {
            var at = asyncJs.LastIndexOf('}');
            await Edit("action", asyncJs.Insert(at, before + "\n"));
            await test(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = at + before.Length, ["typeText"] = "." });
            await CheckWidget("suggestions", "JavaScript Action dot completion offers " + string.Join(", ", expected), expected);
            await test(new Dictionary<string, object?> { ["action"] = "assistance", ["command"] = "hideSuggestWidget" });
        }
        await Edit("action", actionSource);

        const string ts = "export async function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host) { " +
            "const dir = await host.fs.getUserSelectedDirectory(); if (!dir) return; " +
            "const file = await dir.getFile('report'); const stream = await file.open(); " +
            "const n: number = stream.read(new Uint8Array(16), 0, 16); await stream.close(); " +
            "const [selected] = await host.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'rw'}); " +
            "const book = await selected.openWorkbook(); await book.setValues('Sheet1','A1',[[42]]); " +
            "await workbook.saveAs(file,'xlsx'); await book.saveAs(file,'xlsq'); await book.close(); }";
        await Edit("main", ts);
        foreach (var (expression, member) in new[] { ("host.", "fs"), ("dir.", "getDirectory"), ("stream.", "readLine"),
            ("selected.", "openWorkbook"), ("await book.", "getWorksheetNames"), ("await book.", "getUsedRange"),
            ("await book.", "getValues"), ("await book.", "setValues"), ("await book.", "getFormulas"), ("await book.", "setFormulas"),
            ("await workbook.", "saveAs"), ("await book.", "saveAs") })
        {
            var state = await test(new Dictionary<string, object?> { ["action"] = "inspect", ["offset"] = ts.IndexOf(expression, StringComparison.Ordinal) + expression.Length });
            check(state.GetProperty("diagnostics").GetArrayLength() == 0 &&
                state.GetProperty("completions").EnumerateArray().Any(c => c.GetString() == member),
                $"The TypeScript worker resolves {expression}{member} from the v4 host.fs declarations");
        }
        await test(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = ts.IndexOf("stream.read(", StringComparison.Ordinal) + "stream.read(".Length,
            ["button"] = "editor.action.triggerParameterHints" });
        await CheckWidget("parameters", "The actual parameter hint for the buffered file read shows its byte array, offset and length",
            "Uint8Array", "offset", "len");
        await test(new Dictionary<string, object?> { ["action"] = "assistance", ["command"] = "closeParameterHints" });
        await Edit("main", mainSource);
    }
}
