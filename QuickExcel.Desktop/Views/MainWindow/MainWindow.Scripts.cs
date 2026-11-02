using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private string _selectedModule = "main";
    private sealed record DiagnosticEntry(string Label, int Sheet = -1, int Row = -1, int Column = -1, string? Module = null, int Line = 0);

    private void LoadScriptModules()
    {
        var modules = _document.Snapshot.ScriptModules;
        _selectedModule = modules.EntryModules[0];
        ModulePicker.ItemsSource = modules.EntryModules.Concat(modules.Modules.Keys.Order(StringComparer.Ordinal)).Distinct().ToArray();
        ModulePicker.SelectedItem = _selectedModule;
        ScriptEditor.Text = modules.Modules[_selectedModule].Source;
        ModuleTitle.Text = modules.Modules[_selectedModule].Path;
        RefreshModuleRole();
        ScriptPosition.Text = "1:1";
        DiagnosticsList.ItemsSource = null;
        RefreshSignatureState();
    }

    private void CommitScriptModule()
    {
        var modules = _document.Snapshot.ScriptModules;
        if (!modules.Modules.ContainsKey(_selectedModule)) _selectedModule = modules.EntryModules[0];
        // Saving/navigation also flushes editors. An unchanged source is not an edit,
        // including in a read-only workbook whose session rejects every command.
        if (modules.Modules[_selectedModule].Source == ScriptEditor.Text) return;
        if (_selectedModule == modules.EntryModules[0]) _editor.UpdateScriptSource(ScriptEditor.Text);
        else _editor.Execute(new UpdateScriptModulesCommand(new ScriptModuleSet(
            modules.Modules.SetItem(_selectedModule, modules.Modules[_selectedModule] with { Source = ScriptEditor.Text }),
            modules.EntryModules, modules.Revision, modules.BindingRevision, modules.ScriptApiVersion, modules.Bindings)));
    }

    private void ModulePicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_refreshing || _document is null || ModulePicker.SelectedItem is not string id || id == _selectedModule) return;
        try
        {
            CommitScriptModule();
            _refreshing = true;
            _selectedModule = id;
            ScriptEditor.Text = _document.Snapshot.ScriptModules.Modules[id].Source;
            ScriptEditor.IsUndoEnabled = false; ScriptEditor.IsUndoEnabled = true;
            ModuleTitle.Text = _document.Snapshot.ScriptModules.Modules[id].Path;
            RefreshModuleRole();
        }
        catch (Exception ex)
        {
            _refreshing = true; ModulePicker.SelectedItem = _selectedModule; ReportEditFailure(ex);
        }
        finally { _refreshing = false; }
    }

    private void AddModule_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var id = PromptText("新建模块", "模块 ID（字母、数字、下划线）", "helpers");
        if (id is null) return;
        try
        {
            var old = _document.Snapshot.ScriptModules;
            if (old.ScriptApiVersion != 2) throw new InvalidOperationException("此工作簿使用脚本 API v1。请使用 API v2 工作簿添加模块，避免改变现有脚本语义。");
            if (old.Modules.ContainsKey(id)) throw new ArgumentException("模块 ID 已存在。");
            _editor.Execute(new UpdateScriptModulesCommand(new ScriptModuleSet(old.Modules.Add(id, new($"scripts/{id}.js", "export function example(value) {\n    return value;\n}\n")),
                old.EntryModules, old.Revision, old.BindingRevision, 2, old.Bindings)));
            _refreshing = true; ModulePicker.ItemsSource = old.EntryModules.Concat(_document.Snapshot.ScriptModules.Modules.Keys.Order()).Distinct().ToArray();
            _refreshing = false; ModulePicker.SelectedItem = id;
            ShowDiagnostics(); UpdateChrome();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { _refreshing = false; }
    }

    private void RefreshModuleRole()
    {
        var modules = _document.Snapshot.ScriptModules;
        ModuleRole.IsEnabled = modules.ScriptApiVersion == 2;
        var formula = modules.Bindings.Any(b => b.Module == _selectedModule && b.Kind == "formula-functions");
        var actions = modules.Bindings.Any(b => b.Module == _selectedModule && b.Kind == "actions");
        ModuleRole.SelectedIndex = (formula ? 1 : 0) + (actions ? 2 : 0);
    }

    private void ModuleRole_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_refreshing || _document is null || _busy || ModuleRole.SelectedIndex < 0) return;
        try
        {
            CommitScriptModule();
            var old = _document.Snapshot.ScriptModules;
            var bindings = old.Bindings.Where(b => b.Module != _selectedModule).ToList();
            if ((ModuleRole.SelectedIndex & 1) != 0) bindings.Add(new(_selectedModule, "formula-functions"));
            if ((ModuleRole.SelectedIndex & 2) != 0) bindings.Add(new(_selectedModule, "actions"));
            var entries = old.EntryModules.Where(id => bindings.Any(b => b.Module == id)).ToList();
            if (bindings.Any(b => b.Module == _selectedModule) && !entries.Contains(_selectedModule)) entries.Add(_selectedModule);
            if (entries.Count == 0) throw new InvalidOperationException("工作簿至少需要一个公式或操作入口。");
            _editor.Execute(new UpdateScriptModulesCommand(new ScriptModuleSet(old.Modules, entries, old.Revision, old.BindingRevision, old.ScriptApiVersion, bindings)));
            ActionPicker.ItemsSource = null; ShowDiagnostics(); UpdateChrome();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { _refreshing = true; RefreshModuleRole(); _refreshing = false; }
    }

    private void ScriptEditor_SelectionChanged(object sender, RoutedEventArgs e)
    {
        if (ScriptPosition is null) return;
        var line = ScriptEditor.GetLineIndexFromCharacterIndex(ScriptEditor.CaretIndex);
        ScriptPosition.Text = $"{line + 1}:{ScriptEditor.CaretIndex - ScriptEditor.GetCharacterIndexFromLineIndex(Math.Max(0, line)) + 1}";
    }

    private void RefreshDiagnostics()
    {
        var entries = new List<DiagnosticEntry>();
        var modules = _document.Snapshot.ScriptModules;
        foreach (var error in _document.ScriptErrors.Take(100))
        {
            var module = modules.Modules.FirstOrDefault(m => error.Contains(m.Value.Path, StringComparison.Ordinal) || error.Contains(m.Key + ":", StringComparison.Ordinal));
            var names = module.Key is not null ? "(?:" + Regex.Escape(module.Value.Path) + "|" + Regex.Escape(module.Key) + ")"
                : "(?:<quickexcel>|quickexcel|<eval>|script)";
            var location = Regex.Match(error, names + @":(\d+)(?::\d+)?(?:\)|\s|$)");
            var line = location.Success && int.TryParse(location.Groups[1].Value, out var parsed) ? parsed : 0;
            entries.Add(new(error, Module: module.Key ?? modules.EntryModules[0], Line: line));
        }
        for (var index = 0; index < _document.Snapshot.Sheets.Length && entries.Count < 200; index++)
        {
            var sheet = _document.Snapshot.Sheets[index];
            foreach (var (position, cell) in sheet.Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
            {
                if (cell.Value.Kind != CellValueKind.Error) continue;
                entries.Add(new($"{sheet.Name}!{ColumnName(position.Column)}{position.Row + 1}  {CellInput.DisplayText(cell.Value)}", index, position.Row, position.Column));
                if (entries.Count == 200) break;
            }
        }
        DiagnosticsList.ItemsSource = entries;
    }

    private void Diagnostic_Activate(object sender, MouseButtonEventArgs e) => ActivateDiagnostic();
    private void Diagnostic_KeyDown(object sender, KeyEventArgs e)
    { if (e.Key == Key.Enter) { e.Handled = true; ActivateDiagnostic(); } }
    private void ActivateDiagnostic()
    {
        if (_busy || !CommitEditors() || DiagnosticsList.SelectedItem is not DiagnosticEntry error) return;
        try
        {
            if (error.Sheet >= 0) NavigateToCell(error.Sheet, error.Row, error.Column);
            else
            {
                ShowWorkbookPage(true); ModulePicker.SelectedItem = error.Module;
                var line = Math.Clamp(error.Line - 1, 0, Math.Max(0, ScriptEditor.LineCount - 1));
                var start = ScriptEditor.GetCharacterIndexFromLineIndex(line);
                ScriptEditor.Focus(); ScriptEditor.Select(Math.Max(0, start), Math.Max(0, ScriptEditor.GetLineLength(line)));
                ScriptEditor.ScrollToLine(line);
                if (error.Line <= 0) StatusText.Text = "此诊断未提供源码行号。";
            }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
