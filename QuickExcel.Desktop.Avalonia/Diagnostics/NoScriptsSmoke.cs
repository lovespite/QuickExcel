using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Runtime;
using QuickExcel.Security;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunNoScriptsSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        async Task Blocked(Func<Task> operation, string label)
        {
            try { await operation(); throw new InvalidDataException("Execution was allowed: " + label); }
            catch (InvalidOperationException ex)
            { Check(ex.Message.Contains(UiText.Get("当前环境已禁用脚本。")) && ex.Message.Contains("No scripts - "), label); }
        }
        async Task SetPreference(bool value)
        {
            _dialogDriver = window =>
            {
                var box = window.GetVisualDescendants().OfType<CheckBox>().Single(c => c.Name == "DisableScripts");
                Check(box.IsChecked == _settings.Value.DisableScripts, "Settings restores the saved checkbox");
                box.IsChecked = value;
                Check(ScriptExecutionPolicy.Current.UserSetting == value, "Checkbox takes effect before Settings closes");
                window.Close(true); return Task.CompletedTask;
            };
            await OpenSettingsAsync(SettingsSection.Scripts);
        }
        try
        {
            var policy = ScriptExecutionPolicy.Current;
            var startup = policy.Status;
            var external = startup.Replace("User setting/", "").Replace("No scripts - User setting", "");
            Check(_noScriptsStatus.Text == startup && _noScriptsStatus.IsVisible == policy.IsDisabled, "Startup policy is visible on Home");
            _workspace.New(); ShowBook(); _calculationTimer.Stop();
            var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
            {
                ["main"] = new("scripts/main.js", "export function main(w){w.getActiveWorksheet().getRange('A1').setValue('initialized');}"),
                ["action"] = new("scripts/action.js", "export function main(w){w.getActiveWorksheet().getRange('B1').setValue('action');}"),
                ["udf"] = new("scripts/udf.js", "export function VALUE(){return 42;}"),
                ["events"] = new("scripts/events.js", "export function onCellClick(c){c.workbook.getActiveWorksheet().getRange('C1').setValue('event');}")
            }, ["main", "action", "udf", "events"], scriptApiVersion: 4,
                bindings: [new("main", "initialization"), new("action", "actions"), new("udf", "formula-functions"), new("events", "events")]);
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(modules));
            _workspace.Editor.SetCell(0, 0, 3, "=JS.udf.VALUE()");
            if (!policy.IsDisabled)
            {
                await ApplyScriptsAsync();
                Check(Document.ScriptsEnabled && Document.GetCellValue(0, 0, 0) == "initialized" && Document.GetCellValue(0, 0, 3) == "42", "Enabled baseline runs initialization and UDF workers");
                await RunScriptActionAsync(_availableScriptActions.Single());
                Check(Document.GetCellValue(0, 0, 1) == "action", "Enabled baseline runs an action");
            }
            await SetPreference(true);
            Check(!Document.ScriptsEnabled && _availableScriptActions.Count == 0, "Disabling revokes the active contexts and discovered actions");
            Check(_noScriptsStatus.IsVisible && _noScriptsStatus.Text == policy.Status && policy.Status.Contains("User setting"), "Persistent status includes user setting and all external sources");
            var reloaded = new DesktopSettings(Path.Combine(_userDataDirectory, "settings.json")); reloaded.Load();
            Check(reloaded.Value.DisableScripts, "Disable preference persists to disk");
            await Blocked(ApplyScriptsAsync, "Manual Apply is explicitly rejected");
            await Blocked(() => RunScriptActionAsync(null), "Manual Run is explicitly rejected before selection checks");
            await Blocked(RunConsoleAsync, "Console is explicitly rejected");
            await Blocked(SignScriptsAsync, "Signing cannot launch a compiler");
            await Blocked(() => Task.Run(() => new ScriptWorkerRuntimeFactory().Create("throw Error('must not run');")), "Runtime factory cannot bypass desktop policy");
            _automationRun!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            try { await _operation; } catch (InvalidOperationException) { }
            Check(_status.Text!.Contains(UiText.Get("当前环境已禁用脚本。")), "Clicking Run displays the disabled-environment message");

            // Trust and source signing are data operations; they must never override the execution veto.
            using var certificate = ScriptSignatures.CreateSelfSigned("No scripts isolated smoke");
            modules = Document.Snapshot.ScriptModules;
            _workspace.Editor.Execute(new UpdateScriptSignatureCommand(modules, ScriptSignatures.Sign(modules, certificate)));
            ScriptTrust.Trust(Document.Snapshot.ScriptModules); ScriptTrust.AcceptExecutionPolicy(Document.Snapshot.ScriptModules);
            Check(ScriptTrust.Verify(Document.Snapshot.ScriptModules) is { CanAutoEnable: true, ExecutionPolicyAccepted: true }, "Fixture has a valid trusted signature and accepted engine policy");
            var path = Path.Combine(directory, "trusted.xlsq");
            await _workspace.SaveAsync(new(path));
            await OpenPathAsync(path);
            Check(!Document.ScriptsEnabled && _availableScriptActions.Count == 0, "Reopening a fully trusted workbook does not auto-enable scripts");
            var revision = Document.Snapshot.Revision;
            Document.Session.RaiseScriptEvent("cell.click", Document.Snapshot.ActiveSheetId);
            await DispatchEventsAsync();
            await Task.Run(() => Document.Calculate());
            Check(Document.Snapshot.Revision == revision && !Document.ScriptsEnabled &&
                Document.Snapshot.Sheets[0].GetCell(0, 3)!.Calculation?.State == CalculationState.Blocked,
                "Events and recalculation cannot execute disabled scripts; retained UDF caches are marked blocked");
            await SetPreference(false);
            Check(policy.Status == external && _noScriptsStatus.IsVisible == policy.IsDisabled && !Document.ScriptsEnabled, "Clearing user setting retains external vetoes and never silently reauthorizes");
            if (policy.IsDisabled) await Blocked(ApplyScriptsAsync, "External veto survives clearing the checkbox");
            else
            {
                await ApplyScriptsAsync();
                Check(Document.ScriptsEnabled && Document.GetCellValue(0, 0, 3) == "42", "Explicit Apply can authorize fresh contexts after re-enabling");
            }
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
