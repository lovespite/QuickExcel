using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private Task _gridInput = Task.CompletedTask;
    private bool _gridEditorArmed => !_grid.IsEditing;
    private void InstallInputHandlers()
    {
        AddHandler(TextCompositionManager.PreviewTextInputStartEvent, new TextCompositionEventHandler(CompositionStarted), true);
        AddHandler(TextCompositionManager.PreviewTextInputUpdateEvent, new TextCompositionEventHandler(CompositionStarted), true);
        AddHandler(TextCompositionManager.PreviewTextInputEvent, new TextCompositionEventHandler(CompositionCompleted), true);
        PreviewMouseDown += (_, e) =>
        {
            if (_busy)
            {
                e.Handled = e.OriginalSource is not DependencyObject busyTarget ||
                    !(ReferenceEquals(busyTarget, StopCalculationButton) || StopCalculationButton.IsAncestorOf(busyTarget));
                return;
            }
            // Focus loss makes Windows Pinyin commit its unfinished Latin buffer. Keep focus
            // in the composing cell editor until the user confirms/cancels the candidate.
            if (_editState.IsComposing && Keyboard.FocusedElement is TextBox editor && IsCellInputSource(editor) &&
                e.OriginalSource is DependencyObject target && !ReferenceEquals(target, editor) && !editor.IsAncestorOf(target))
            {
                e.Handled = true;
                CanCommitComposition();
            }
        };
        PreviewMouseWheel += (_, e) => { if (_busy) e.Handled = true; };
        CommandManager.AddPreviewExecutedHandler(this, (_, e) => { if (_busy) e.Handled = true; });
        SheetGrid.PreviewKeyDown += Grid_KeyDown;
        // While a text editor owns focus, its Ctrl+Z/C/Y/V remain normal text commands.
        SheetGrid.CommandBindings.Add(new CommandBinding(ApplicationCommands.Copy,
            (_, e) =>
            {
                if (_editState.GridEditing) return;
                e.Handled = true;
                try { CopySelection(); }
                catch (Exception ex) { ReportEditFailure(ex); }
            },
            (_, e) => { if (!_editState.GridEditing) { e.CanExecute = !_busy; e.Handled = true; } }));
    }

    private bool IsCellInputSource(object source) => source is DependencyObject obj &&
        (ReferenceEquals(obj, FormulaBar) || FormulaBar.IsAncestorOf(obj) || SheetGrid.IsAncestorOf(obj));

    private void CompositionStarted(object sender, TextCompositionEventArgs e)
    {
        if (_busy) { e.Handled = true; return; }
        if (IsCellInputSource(e.OriginalSource))
        {
            ActivateGridDraft(e.OriginalSource);
            _editState.StartComposition();
        }
    }

    private void CompositionCompleted(object sender, TextCompositionEventArgs e)
    {
        if (_busy) { e.Handled = true; return; }
        if (!IsCellInputSource(e.OriginalSource)) return;
        ActivateGridDraft(e.OriginalSource);
        // Preview runs before TextBox inserts the final composed text. Keep the gate until bubbling completes.
        Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(_editState.EndComposition));
    }

    private void ActivateGridDraft(object source)
    {
        if (ReferenceEquals(source, _grid.Editor)) _grid.ActivateDraft();
    }

    private bool CanCommitComposition()
    {
        if (!_editState.IsComposing && !_grid.IsComposing) return true;
        StatusText.Text = "请先完成或取消输入法组合输入，再提交、切换或保存。";
        return false;
    }

    private async void FormulaBar_KeyDown(object sender, KeyEventArgs e)
    {
        if (_editState.IsComposing || e.Key == Key.ImeProcessed) return;
        if (e.Key == Key.Escape)
        {
            e.Handled = true;
            _editState.EndFormula();
            UpdateFormulaBar();
            GridSurface.Select(new(_selectedRow, _selectedColumn));
            StatusText.Text = "已取消公式栏编辑。";
        }
        else if (e.Key is Key.Enter or Key.Tab)
        {
            e.Handled = true;
            await NavigateGridAsync(e.Key, Keyboard.Modifiers);
        }
    }

    private async void Grid_KeyDown(object sender, KeyEventArgs e)
    {
        if (_busy || _editState.IsComposing || e.Key == Key.ImeProcessed) return;
        var modifiers = Keyboard.Modifiers;
        var editing = !_gridEditorArmed && (e.OriginalSource is TextBox || _editState.GridEditing);
        var ctrl = modifiers.HasFlag(ModifierKeys.Control);
        if (modifiers.HasFlag(ModifierKeys.Alt))
        {
            if ((e.Key == Key.Down || e.SystemKey == Key.Down) && CurrentView.Region is not null)
            {
                e.Handled = true;
                var target = _grid.Editor;
                await (_workspaceInput = OpenHeaderFilterAsync(_selectedColumn, target));
            }
            return;
        }
        var navigation = e.Key is Key.Enter or Key.Tab or Key.Left or Key.Right or Key.Up or Key.Down or Key.Home or Key.End;
        var handled = editing
            ? !ctrl && e.Key is Key.Enter or Key.Tab or Key.Escape
            : navigation || e.Key is Key.F2 or Key.Delete or Key.Escape || ctrl && e.Key is Key.C or Key.X or Key.V or Key.A;
        if (handled)
        {
            e.Handled = true;
            try { await (_gridInput = HandleGridKeyAsync(e.Key, modifiers, editing)); }
            catch (Exception ex) { ReportEditFailure(ex); }
        }
    }

    private async Task HandleGridKeyAsync(Key key, ModifierKeys modifiers, bool editing)
    {
        if (_busy || !CanCommitComposition()) return;
        var ctrl = modifiers.HasFlag(ModifierKeys.Control);
        if (key == Key.Escape)
        {
            _workbookClipboard = null;
            GridSurface.Cancel();
            _editState.EndGrid();
            UpdateFormulaBar();
            StatusText.Text = "已取消单元格编辑及内部复制/剪切。";
            return;
        }
        if (ctrl && key == Key.C) { CopySelection(); return; }
        if (ctrl && key == Key.X) { CopySelection(cut: true); return; }
        if (ctrl && key == Key.V)
        {
            await PasteClipboardAsync(modifiers.HasFlag(ModifierKeys.Shift) ? "formulas" : "all");
            return;
        }
        if (ctrl && key == Key.A)
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex];
            GridSurface.SelectRange(new(0, 0, Math.Min(_rowCount, Math.Max(1, sheet.UsedRowCount)) - 1,
                Math.Min(_columnCount, Math.Max(1, sheet.UsedColumnCount)) - 1), new(0, 0));
            return;
        }
        if (key == Key.F2)
        {
            if (!CommitFormulaBar()) return;
            GridSurface.BeginEdit();
            _grid.Editor.Select(_grid.Editor.Text.Length, 0);
            return;
        }
        if (key == Key.Delete)
        {
            var selected = GridSurface.Selection;
            if (selected.Count > GridClipboard.MaxEditCells) throw new ArgumentException("单次清空上限为 20,000 格。");
            if (!CommitEditors()) return;
            CurrentView.RequireCurrentFilter(_document.Snapshot);
            var result = _editor.ClearCells(_sheetIndex, selected.Where(p => !CurrentView.IsFilteredOut(_document.Snapshot, _sheetIndex, p.Row)).Select(p => (p.Row, p.Column)));
            _calculationPending |= result.Changed;
            UpdateChrome();
            if (_calculationPending) await RecalculateAsync();
            return;
        }
        await NavigateGridAsync(key, modifiers);
    }

    private async Task NavigateGridAsync(Key key, ModifierKeys modifiers)
    {
        if (!CommitEditors()) return;
        var shift = modifiers.HasFlag(ModifierKeys.Shift);
        var ctrl = modifiers.HasFlag(ModifierKeys.Control);
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var row = _selectedRow;
        var column = _selectedColumn;
        var currentRange = _grid.CurrentRange;
        var delta = shift && key is Key.Tab or Key.Enter ? -1 : 1;
        switch (key)
        {
            case Key.Enter: row = delta > 0 ? currentRange.Bottom + 1 : currentRange.Top - 1; break;
            case Key.Tab:
                column = delta > 0 ? currentRange.Right + 1 : currentRange.Left - 1;
                if (column >= _columnCount) { column = 0; row++; }
                if (column < 0) { column = _columnCount - 1; row--; }
                break;
            case Key.Left: column = ctrl ? 0 : currentRange.Left - 1; break;
            case Key.Right: column = ctrl ? Math.Max(0, sheet.UsedColumnCount - 1) : currentRange.Right + 1; break;
            case Key.Up: row = ctrl ? 0 : currentRange.Top - 1; break;
            case Key.Down: row = ctrl ? Math.Max(0, sheet.UsedRowCount - 1) : currentRange.Bottom + 1; break;
            case Key.Home: column = 0; if (ctrl) row = 0; break;
            case Key.End: column = Math.Max(0, sheet.UsedColumnCount - 1); if (ctrl) row = Math.Max(0, sheet.UsedRowCount - 1); break;
        }
        row = Math.Clamp(row, 0, _rowCount - 1);
        column = Math.Clamp(column, 0, _columnCount - 1);
        // Imported hidden dimensions remain hidden, and ordinary navigation skips them.
        var rowDirection = row < _selectedRow ? -1 : 1;
        while ((sheet.Rows.GetValueOrDefault(row)?.Hidden == true || CurrentView.IsFilteredOut(_document.Snapshot, _sheetIndex, row)) && row + rowDirection >= 0 && row + rowDirection < _rowCount) row += rowDirection;
        var columnDirection = column < _selectedColumn ? -1 : 1;
        while (sheet.Columns.TryGetValue(column, out var cm) && cm.Hidden && column + columnDirection >= 0 && column + columnDirection < _columnCount) column += columnDirection;
        GridSurface.Select(new(row, column), shift && key is not (Key.Tab or Key.Enter));
        UpdateFormulaBar();
        if (_calculationPending) await RecalculateAsync();
    }

    private void CopySelection(bool cut = false)
    {
        if (!CommitEditors()) return;
        CurrentView.RequireCurrentFilter(_document.Snapshot);
        bool Excluded(int row) => CurrentView.IsFilteredOut(_document.Snapshot, _sheetIndex, row) ||
            _document.Snapshot.Sheets[_sheetIndex].Rows.GetValueOrDefault(row)?.Hidden == true;
        var range = GridRange.Rectangle(GridSurface.Selection, Excluded);
        var text = GridClipboard.Write(range, (r, c) => _document.GetCellValue(_sheetIndex, r, c), Excluded);
        if (Enumerable.Range(range.Top, range.Rows).Any(Excluded) || range.Count > QuickExcel.Core.WorkbookClipboard.MaxCells)
        {
            if (cut) throw new InvalidOperationException("剪切需要不含隐藏行且不超过 20,000 格的连续矩形。");
            _workbookClipboard = null; Clipboard.SetText(text);
            StatusText.Text = "已仅复制可见行的显示文本；内部复制需要不含隐藏行且不超过 20,000 格的矩形。";
            return;
        }
        CaptureWorkbookClipboard(range, text, cut);
    }

    private async Task PasteSelectionAsync(string text, bool allowFormulas)
    {
        var selection = GridSurface.Selection;
        var range = selection.Count == 0 ? new GridRange(_selectedRow, _selectedColumn, _selectedRow, _selectedColumn) : GridRange.Rectangle(selection);
        // A single merged cell is one paste destination even though its selection spans all covered cells.
        var merge = _grid.CurrentRange;
        if (merge.Count > 1 && range == new GridRange(merge.Top, merge.Left, merge.Bottom, merge.Right) &&
            GridClipboard.Read(text) is [{ Length: 1 }])
            range = new(merge.Top, merge.Left, merge.Top, merge.Left);
        var edits = GridClipboard.Prepare(text, range, allowFormulas);
        if (!CommitEditors()) return;
        CurrentView.RequireVisibleTargets(_document.Snapshot, _sheetIndex, edits.Select(edit => edit.Row));
        ApplyValues(edits);
        RefreshSheet(true);
        GridSurface.SelectRange(new(range.Top, range.Left, edits.Max(e => e.Row), edits.Max(e => e.Column)), new(range.Top, range.Left));
        if (_calculationPending) await RecalculateAsync();
        StatusText.Text = allowFormulas ? "已按输入解析粘贴（包括公式），可一次撤销。" : "已粘贴值；公式文本保持为文本，可一次撤销。";
    }
}
