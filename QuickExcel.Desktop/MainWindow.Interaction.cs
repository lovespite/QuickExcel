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
    private bool _gridEditorArmed;
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
        SheetGrid.GotKeyboardFocus += (_, e) =>
        {
            if (e.NewFocus is not DataGridCell cell) return;
            // A focused editor must already be registered with TSF before the first native IME
            // key arrives. Preparing it from that key is too late for modern Windows Pinyin.
            Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
            {
                if (_busy || _editState.GridEditing || _editState.IsComposing || !cell.IsKeyboardFocused) return;
                if (!GridSurface.BeginEdit()) return;
                if (cell.Content is not TextBox editor) return;
                GridInputColumn.InitializeText(editor, _document.GetCellValue(_sheetIndex, _selectedRow, _selectedColumn));
                editor.SelectAll();
                GridInputColumn.SetPrepared(editor, true);
                editor.TextChanged += (_, _) => ActivateGridDraft(editor);
                _gridEditorArmed = true;
                _editState.EndGrid(); // Ready for input; no draft yet. Navigation still belongs to the grid.
            }));
        };
        SheetGrid.PreviewMouseLeftButtonDown += Grid_MouseLeftButtonDown;
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

    private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_busy || _editState.IsComposing) return;
        var modifiers = Keyboard.Modifiers;
        if (modifiers.HasFlag(ModifierKeys.Shift)) _grid.CaptureAnchor();
        else _grid.ResetAnchor();

        var source = e.OriginalSource as DependencyObject;
        while (source is not null and not DataGridCell)
            source = source is Visual ? VisualTreeHelper.GetParent(source) : LogicalTreeHelper.GetParent(source);
        if (!_gridEditorArmed || modifiers != ModifierKeys.None || source is not DataGridCell { IsSelected: true, Content: TextBox editor } cell ||
            !ReferenceEquals(editor, SheetGrid.CurrentColumn?.GetCellContent(SheetGrid.CurrentItem))) return;

        if (SheetGrid.SelectedCells.Count == 1)
        {
            BeginExplicitGridEdit(editor);
            editor.Focus();
            editor.CaretIndex = Math.Max(0, editor.GetCharacterIndexFromPoint(e.GetPosition(editor), true));
        }
        else if (cell.DataContext is SheetRow row)
        {
            // Clicking the active cell of a range first reduces it to one selected cell.
            GridSurface.Select(new(row.Index, cell.Column.DisplayIndex));
        }
        e.Handled = true;
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
        if (_gridEditorArmed && source is DependencyObject obj && SheetGrid.IsAncestorOf(obj))
        {
            _gridEditorArmed = false;
            if (SheetGrid.CurrentColumn?.GetCellContent(SheetGrid.CurrentItem) is TextBox editor)
                GridInputColumn.SetPrepared(editor, false);
            _editState.BeginGrid();
        }
    }

    private void BeginExplicitGridEdit(TextBox editor)
    {
        _gridEditorArmed = false;
        GridInputColumn.SetPrepared(editor, false);
        _editState.BeginGrid();
        GridInputColumn.InitializeText(editor, _document.GetCellInput(_sheetIndex, _selectedRow, _selectedColumn));
    }


    private bool CanCommitComposition()
    {
        if (!_editState.IsComposing) return true;
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
        if (modifiers.HasFlag(ModifierKeys.Alt)) return;
        var navigation = e.Key is Key.Enter or Key.Tab or Key.Left or Key.Right or Key.Up or Key.Down or Key.Home or Key.End;
        var handled = editing
            ? !ctrl && e.Key is Key.Enter or Key.Tab or Key.Escape
            : navigation || e.Key is Key.F2 or Key.Delete || ctrl && e.Key is Key.C or Key.V or Key.A;
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
            GridSurface.Cancel();
            _editState.EndGrid();
            UpdateFormulaBar();
            StatusText.Text = "已取消单元格编辑。";
            return;
        }
        if (ctrl && key == Key.C) { CopySelection(); return; }
        if (ctrl && key == Key.V)
        {
            if (Clipboard.ContainsText()) await PasteSelectionAsync(Clipboard.GetText(), modifiers.HasFlag(ModifierKeys.Shift));
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
            if (SheetGrid.CurrentColumn?.GetCellContent(SheetGrid.CurrentItem) is TextBox editor)
            {
                if (_gridEditorArmed) BeginExplicitGridEdit(editor);
                editor.Select(editor.Text.Length, 0);
            }
            return;
        }
        if (key == Key.Delete)
        {
            var selected = GridSurface.Selection;
            if (selected.Count > GridClipboard.MaxEditCells) throw new ArgumentException("单次清空上限为 20,000 格。");
            if (!CommitEditors()) return;
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
        var delta = shift && key is Key.Tab or Key.Enter ? -1 : 1;
        switch (key)
        {
            case Key.Enter: row += delta; break;
            case Key.Tab:
                column += delta;
                if (column >= _columnCount) { column = 0; row++; }
                if (column < 0) { column = _columnCount - 1; row--; }
                break;
            case Key.Left: column = ctrl ? 0 : column - 1; break;
            case Key.Right: column = ctrl ? Math.Max(0, sheet.UsedColumnCount - 1) : column + 1; break;
            case Key.Up: row = ctrl ? 0 : row - 1; break;
            case Key.Down: row = ctrl ? Math.Max(0, sheet.UsedRowCount - 1) : row + 1; break;
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

    private void CopySelection()
    {
        if (!CommitEditors()) return;
        bool Excluded(int row) => CurrentView.IsFilteredOut(_document.Snapshot, _sheetIndex, row);
        var range = GridRange.Rectangle(GridSurface.Selection, Excluded);
        var text = GridClipboard.Write(range, (r, c) => _document.GetCellValue(_sheetIndex, r, c), Excluded);
        Clipboard.SetText(text);
        StatusText.Text = $"已复制 {range.Rows} 行 × {range.Columns} 列显示结果。";
    }

    private async Task PasteSelectionAsync(string text, bool allowFormulas)
    {
        var selection = GridSurface.Selection;
        var range = selection.Count == 0 ? new GridRange(_selectedRow, _selectedColumn, _selectedRow, _selectedColumn) : GridRange.Rectangle(selection);
        var edits = GridClipboard.Prepare(text, range, allowFormulas);
        CurrentView.RequireVisibleTargets(_document.Snapshot, _sheetIndex, edits.Select(edit => edit.Row));
        if (!CommitEditors()) return;
        ApplyValues(edits);
        RefreshSheet(true);
        GridSurface.SelectRange(new(range.Top, range.Left, edits.Max(e => e.Row), edits.Max(e => e.Column)), new(range.Top, range.Left));
        if (_calculationPending) await RecalculateAsync();
        StatusText.Text = allowFormulas ? "已按输入解析粘贴（包括公式），可一次撤销。" : "已粘贴值；公式文本保持为文本，可一次撤销。";
    }
}
