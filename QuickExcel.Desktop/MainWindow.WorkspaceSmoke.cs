using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunWorkspaceSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        var cases = new List<string>();
        void Check(bool condition, string name) { if (!condition) throw new Exception("M5 workspace: " + name); cases.Add(name); }
        async Task Idle() => await Dispatcher.InvokeAsync(() => { UpdateLayout(); }, DispatcherPriority.ApplicationIdle);
        void Structure(string operation, string? answer = null)
        { _smokeTextDecision = answer; Structure_Click(new MenuItem { Tag = operation }, new()); _smokeTextDecision = null; }
        void Style(string operation) => Style_Click(new MenuItem { Tag = operation }, new());
        var path = Path.Combine(folder, "workspace.xlsq");
        _suppressAutoCalculation = true;
        try
        {
            await RunCalculationRefreshSmokeAsync(Check);
            await RunGridResizeSmokeAsync(folder, Check);
            await RunAutoFitSmokeAsync(folder, Check);
            InstallDocument(WorkbookDocument.CreateBlank(_services), path, "工作区验收.xlsq");
            ApplyValues([(0, 0, "Key"), (0, 1, "Label"), (0, 2, "Value"), (1, 0, "2"), (1, 1, "second"), (1, 2, "=A2*10"),
                (2, 0, "1"), (2, 1, "first"), (2, 2, "=A3*10"), (3, 0, "2"), (3, 1, "third"), (3, 2, "=A4*10")]);
            RefreshSheet(true); await RecalculateAsync(); await Idle();
            GridSurface.SelectRange(new(1, 0, 3, 2), new(1, 0));
            Sort_Click(new MenuItem { Tag = "asc" }, new()); await _workspaceInput;
            await RecalculateAsync();
            Check(_document.GetCellValue(0, 1, 1) == "first" && _document.GetCellValue(0, 2, 1) == "second" && _document.GetCellValue(0, 3, 1) == "third", "Sort menu uses stable rectangular row order");
            Check(_document.GetCellValue(0, 1, 2) == "10" && _document.GetCellValue(0, 2, 2) == "20", "Sorted relative formulas calculate at their new rows");
            Check(_selectedRow == 1 && _selectedColumn == 0 && CellAddress.Text == "A2", "Sort restores current address and focus target");
            await UndoAsync(false); Check(_document.GetCellValue(0, 1, 1) == "second", "One undo restores entire sort");
            await UndoAsync(true); Check(_document.GetCellValue(0, 1, 1) == "first", "Redo reapplies sort without invoking scripts");
            GridSurface.Select(new(1, 0)); Structure("insertRow"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 2, 1) == "first" && _document.GetCellInput(0, 2, 2) == "=(A3*10)", "Insert row menu updates formula targets");
            Structure("deleteRow"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 1, 1) == "first", "Delete row menu restores positions");
            GridSurface.Select(new(1, 0)); Structure("moveRow", "4"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 3, 1) == "first" && _selectedRow == 3, "Move row dialog relocates data and selection");
            await UndoAsync(false);
            GridSurface.Select(new(1, 1)); Structure("insertColumn"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 1, 2) == "first", "Insert column menu moves contents");
            Structure("deleteColumn"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 1, 1) == "first", "Delete column menu moves contents back");
            GridSurface.Select(new(1, 1)); Structure("moveColumn", "4"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 1, 3) == "first" && _selectedColumn == 3, "Move column dialog relocates selection");
            await UndoAsync(false);
            Structure("rename", "数据 工作表"); await RecalculateAsync();
            Check(SheetPicker.SelectedItem?.ToString() == "数据 工作表", "Rename updates sheet picker");
            GridSurface.Select(new(0, 0)); Structure("sizeRow", "42"); Structure("sizeColumn", "24"); await Idle();
            var resized = _grid.Control.GetModelCellBounds(0, 0);
            Check(resized is { } size && Math.Abs(size.Height - 42) < 1 && Math.Abs(size.Width - 173) < 1,
                $"Row height and column width commands reach pixel-rounded geometry: {resized}");
            GridSurface.Select(new(0, 3)); Structure("hideColumn");
            Check(_grid.Control.GetModelCellBounds(0, 3) is null, "Hide column applies immediately");
            Structure("showColumn"); Check(_grid.Control.GetModelCellBounds(0, 3) is not null,
                $"Show column applies immediately: hidden={_grid.Sheet.Columns.GetValueOrDefault(3)?.Hidden}, current={_grid.Current}, selection={string.Join(';', _grid.SelectionRanges)}, status={StatusText.Text}");
            GridSurface.Select(new(0, 0)); Style("bold"); Style("fill"); Style("border"); Style("center");
            var style = _document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)];
            Check(style.Bold && style.FillColor == "FFFFE599" && style.Horizontal == 2 && style.BottomBorder == 1, "Format menu composes font fill border alignment without losing previous style");
            GridSurface.Select(new(1, 2)); Style("number"); await RecalculateAsync();
            Check(_document.GetCellInput(0, 1, 2) == "=(A2*10)" && _document.GetCellValue(0, 1, 2).Contains("10.00"), "Number format changes display while formula bar retains raw formula");
            ApplyValues([(6, 0, "46000")]); GridSurface.Select(new(6, 0)); Style("date");
            Check(_document.GetCellValue(0, 6, 0) == "2025-12-09" && _document.GetCellInput(0, 6, 0) == "46000", "Date menu uses normalized custom format and leaves raw serial unchanged");
            var revision = _document.Snapshot.Revision;
            ApplyFilter(new(0, 0, 3, 2), 1, "first"); await Idle();
            Check(CurrentView.IsFilteredOut(_document.Snapshot, 0, 2) && _document.Snapshot.Revision == revision, "Filter hides rows without a document revision");
            Check(_grid.Control.GetModelCellBounds(2, 0) is null, "Filter reaches WPF row visibility");
            GridSurface.SelectRange(new(0, 0, 3, 2), new(0, 0)); CopySelection();
            Check(Clipboard.GetText().Contains("first") && !Clipboard.GetText().Contains("second"), "Filtered copy contains only visible rows");
            GridSurface.Select(new(1, 0));
            try { await PasteSelectionAsync("9\n8", false); throw new Exception("Filtered paste accepted hidden target"); }
            catch (InvalidOperationException) { }
            Check(_document.Snapshot.Revision == revision, "Paste spanning a filtered row is rejected atomically");
            Sort_Click(new MenuItem { Tag = "asc" }, new()); await _workspaceInput;
            Check(_document.Snapshot.Revision == revision && StatusText.Text.Contains("清除筛选"), "Sort under filter is explicitly refused");
            GridSurface.Select(new(1, 1)); FormulaBar.Text = "no longer matches"; await ApplyFormulaBarAsync();
            Check(_selectedRow == 0 && CellAddress.Text == "B1", "Editing a filter key moves focus to visible header when the row becomes hidden");
            await UndoAsync(false); ApplyFilter(new(0, 0, 3, 2), 1, "first");
            GridSurface.SelectRange(new(1, 1, 3, 1), new(1, 1));
            await HandleGridKeyAsync(Key.Delete, ModifierKeys.None, false);
            Check(_document.GetCellInput(0, 1, 1) == "" && _document.GetCellInput(0, 2, 1) == "second", "Filtered Delete clears visible cells only");
            await UndoAsync(false); Check(CurrentView.FilterRange is null && _document.GetCellValue(0, 1, 1) == "first", "Undo restores cells and clears stale view range");
            GridSurface.Select(new(1, 1)); Freeze_Click(new MenuItem(), new()); await Idle();
            Check(_grid.Control.FrozenRowCount == 1 && _grid.Control.FrozenColumnCount == 1, "Freeze creates native frozen rows and columns");
            GridSurface.Select(new(80, 10)); await Idle();
            Check(_grid.Control.GetModelCellBounds(0, 0) is not null && _grid.Control.FrozenColumnCount == 1, "Frozen row and first column survive distant scrolling");
            GridSurface.Select(new(0, 0)); await Idle();
            Check(_selectedRow == 0 && _selectedColumn == 0, "Selecting frozen row locates original editable cell");
            GridSurface.Select(new(1, 1));
            CaptureGridEvidence(Path.Combine(folder, "QuickExcel-workspace.png"));
            var firstSheet = _document.Snapshot.Sheets[0].Id;
            AddSheet_Click(this, new()); await Idle();
            Check(_grid.Control.FrozenColumnCount == 0, "New sheet starts with independent view state");
            _smokeDiscardDecision = MessageBoxResult.Cancel; Structure("deleteSheet");
            Check(_document.Snapshot.Sheets.Length == 2, "Cancel delete sheet leaves workbook unchanged");
            _smokeDiscardDecision = MessageBoxResult.OK; Structure("deleteSheet");
            Check(_document.Snapshot.Sheets.Length == 1 && _document.Snapshot.Sheets[0].Id == firstSheet, "Delete sheet menu retains surviving sheet identity");
            await UndoAsync(false); Check(_document.Snapshot.Sheets.Length == 2, "Undo restores deleted sheet through WPF history");
            _smokeDiscardDecision = null;
            SheetPicker.SelectedIndex = 0; await Idle();
            GridSurface.Select(new(1, 1)); Freeze_Click(new MenuItem(), new());
            SheetPicker.SelectedIndex = 1; SheetPicker.SelectedIndex = 0; await Idle();
            Check(_grid.Control.FrozenColumnCount == 1 && _document.Snapshot.Sheets[0].Id == firstSheet, "Switching sheet restores its view and stable identity");
            Check(await SaveAsync(false) && !_document.IsDirty && !Title.StartsWith("*"), "Save updates dirty title and stores captured revision");
            var saved = File.ReadAllBytes(path);
            ApplyValues([(5, 0, "local draft")]);
            File.WriteAllText(path, "external replacement");
            Check(!await SaveAsync(false) && _document.IsDirty && File.ReadAllText(path) == "external replacement", "External replacement blocks UI save and retains local draft");
            var copy = Path.Combine(folder, "workspace-copy.xlsq");
            Check(await SaveToPathAsync(copy, false, new(null)) && !_document.IsDirty, "Save to another path resolves conflict without overwriting source");
            File.WriteAllBytes(path, saved);
            await OpenPathAsync(copy);
            Check(_document.GetCellInput(0, 5, 0) == "local draft" && !_document.ScriptsEnabled && CurrentView.FrozenRows == 0, "Reopen restores data and styles with scripts disabled and fresh view");
            Check(_document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 0)].Bold, "Styled workbook survives actual desktop save and reopen");
            Check(_document.GetCellValue(0, 6, 0) == "2025-12-09", "Custom date format survives UI save and reopen");
            var identity = _document.Session.SessionId;
            await OpenPathAsync(Path.Combine(folder, "missing.xlsq"));
            Check(_document.Session.SessionId == identity, "Missing recent file preserves current document");
            RecentFiles_Opened(this, new());
            Check(_recentFiles.Paths[0] == copy && RecentFilesMenu.Items.Count >= 3, "Successful opens and saves populate recent menu; failures do not");
            var reload = new RecentFileStore(Path.Combine(folder, "recent-files.json")); reload.Load();
            Check(reload.Paths.Contains(copy), "Recent files persist to isolated acceptance settings");
            ScriptPage_Click(this, new()); Check(ScriptPanel.Visibility == Visibility.Visible && !WorkbookPage.IsVisible, "Workspace opens dedicated script page");
            WorkbookPage_Click(this, new()); Check(ScriptPanel.Visibility == Visibility.Collapsed && WorkbookPage.IsVisible, "Workspace restores full-width grid");
            ApplyValues([(5, 0, "pending")]); _smokeDiscardDecision = MessageBoxResult.Cancel;
            New_Click(this, new()); await Idle();
            Check(_document.Session.SessionId == identity && _document.IsDirty, "Cancel New preserves dirty workbook");
            _smokeDiscardDecision = MessageBoxResult.No;
            CloseDocument_Click(this, new()); await Idle();
            Check(_filePath is null && _document.Snapshot.Sheets.Length == 1 && _document.GetCellInput(0, 5, 0) == "", "Close workbook with discard opens clean unnamed workspace");
            File.WriteAllText(Path.Combine(folder, "workspace-cases.json"), JsonSerializer.Serialize(new { cases,
                scope = "Actual WPF controls and shared handlers; dialog text/decisions injected. No physical keyboard or system file-dialog automation." }, new JsonSerializerOptions { WriteIndented = true }));
            File.WriteAllText(Path.Combine(folder, "workspace-smoke-result.txt"), $"PASS: {cases.Count} WPF assertions; UI-004 / UI-005.");
        }
        finally { _suppressAutoCalculation = false; _smokeDiscardDecision = null; _smokeTextDecision = null; }
    }
}
