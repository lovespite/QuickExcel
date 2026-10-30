using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private readonly Popup _formulaHelp = new() { Placement = PlacementMode.Bottom, StaysOpen = true, AllowsTransparency = true, Focusable = false };
    private readonly ListBox _formulaSuggestions = new() { Focusable = false, MaxHeight = 210, MinWidth = 280 };
    private readonly TextBlock _formulaHint = new() { Margin = new(6), MaxWidth = 600, TextWrapping = TextWrapping.Wrap };
    private TextBox? _assistEditor;
    private FormulaCompletion? _completion;
    private bool _suppressFormulaHelp;
    private sealed record ReferenceGesture(TextBox Editor, string Original, FormulaTextEdit Slot, CellPosition Start);
    private ReferenceGesture? _referenceGesture;
    private bool _endNavigation;

    private void InstallFormulaEditing()
    {
        PreviewKeyDown += (_, e) =>
        {
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (!e.Handled && key == Key.OemPlus && Keyboard.Modifiers == ModifierKeys.Alt && Keyboard.FocusedElement is { } focus && IsCellInputSource(focus))
            { e.Handled = true; BeginAutoSum(); }
        };
        var stack = new StackPanel(); stack.Children.Add(_formulaSuggestions); stack.Children.Add(_formulaHint);
        _formulaHelp.Child = new Border { Background = Brushes.White, BorderBrush = Brushes.SeaGreen, BorderThickness = new(1), Child = stack };
        foreach (var editor in new[] { FormulaBar, _grid.Editor })
        {
            editor.TextChanged += (_, _) => UpdateFormulaHelp(editor);
            editor.SelectionChanged += (_, _) => UpdateFormulaHelp(editor);
            editor.GotKeyboardFocus += (_, _) => UpdateFormulaHelp(editor);
            editor.LostKeyboardFocus += (_, _) => Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
            { if (Keyboard.FocusedElement is not TextBox focused || focused != FormulaBar && focused != _grid.Editor) CloseFormulaHelp(); }));
        }
        _grid.EditingChanged += () => { if (!_grid.IsEditing) CloseFormulaHelp(); };
        _formulaSuggestions.PreviewMouseDown += (_, e) =>
        {
            if (ItemsControl.ContainerFromElement(_formulaSuggestions, e.OriginalSource as DependencyObject) is ListBoxItem item)
            { _formulaSuggestions.SelectedItem = item.Content; AcceptCompletion(); e.Handled = true; }
        };
        SheetGrid.PreviewMouseDown += (_, e) =>
        {
            if (e.Handled || e.ChangedButton != MouseButton.Left || e.OriginalSource is not DependencyObject target ||
                !ReferenceEquals(target, _grid.Control) && !_grid.Control.IsAncestorOf(target)) return;
            var hit = _grid.Control.HitTestModel(e.GetPosition(_grid.Control));
            if (hit.IsCell && BeginReferenceGesture(new(hit.Row!.Value, hit.Column!.Value)))
            { e.Handled = true; _grid.Control.CaptureMouse(); }
        };
        SheetGrid.PreviewMouseMove += (_, e) =>
        {
            if (_referenceGesture is null) return;
            var hit = _grid.Control.HitTestModel(e.GetPosition(_grid.Control));
            if (hit.IsCell) UpdateReferenceGesture(new(hit.Row!.Value, hit.Column!.Value));
            e.Handled = true;
        };
        SheetGrid.PreviewMouseUp += (_, e) => { if (_referenceGesture is not null && e.ChangedButton == MouseButton.Left) { EndReferenceGesture(); e.Handled = true; } };
        _grid.Control.LostMouseCapture += (_, _) => EndReferenceGesture();
        Deactivated += (_, _) => { EndReferenceGesture(); CloseFormulaHelp(); };
        Closed += (_, _) => _formulaHelp.IsOpen = false;
    }

    private bool IsFormulaDraft(TextBox editor) => !_busy && !_editState.IsComposing && !_grid.IsComposing && editor.Text.StartsWith('=') &&
        (editor == FormulaBar ? _formulaPending : _grid.IsEditing);
    private void CloseFormulaHelp() { _formulaHelp.IsOpen = false; _completion = null; _assistEditor = null; }
    private void UpdateFormulaHelp(TextBox editor)
    {
        if (_suppressFormulaHelp || _referenceGesture is not null) return;
        if (!editor.IsKeyboardFocused || !IsFormulaDraft(editor)) { if (_assistEditor == editor) CloseFormulaHelp(); return; }
        _assistEditor = editor;
        _completion = editor.SelectionLength == 0 ? FormulaEditing.Complete(editor.Text, editor.CaretIndex) : null;
        _formulaSuggestions.ItemsSource = _completion?.Names;
        _formulaSuggestions.SelectedIndex = _completion is null ? -1 : 0;
        _formulaSuggestions.Visibility = _completion is null ? Visibility.Collapsed : Visibility.Visible;
        var hint = FormulaEditing.Hint(editor.Text, editor.CaretIndex);
        _formulaHint.Text = hint is null ? (_completion is null ? "" : "↑↓ 选择 · Tab 补全 · F4 切换引用") : $"{hint.Signature}\n当前第 {hint.Argument} 个参数 · F4 切换引用";
        _formulaHelp.PlacementTarget = editor;
        _formulaHelp.IsOpen = _completion is not null || hint is not null;
    }
    private void ReplaceDraft(TextBox editor, FormulaTextEdit edit, bool select)
    {
        _suppressFormulaHelp = true;
        try
        {
            editor.Select(edit.Start, edit.Length); editor.SelectedText = edit.Text;
            editor.Select(select ? edit.Start : edit.Start + edit.Text.Length, select ? edit.Text.Length : 0);
        }
        finally { _suppressFormulaHelp = false; }
        editor.Focus(); UpdateFormulaHelp(editor);
    }
    private bool AcceptCompletion()
    {
        if (_assistEditor is not { } editor || _completion is not { } completion || _formulaSuggestions.SelectedItem is not string name || !IsFormulaDraft(editor)) return false;
        var after = completion.Start + completion.Length;
        while (after < editor.Text.Length && char.IsWhiteSpace(editor.Text[after])) after++;
        var suffix = after < editor.Text.Length && editor.Text[after] == '(' ? "" : "(";
        ReplaceDraft(editor, new(completion.Start, completion.Length, name + suffix), false);
        if (suffix.Length == 0) editor.CaretIndex = after + name.Length - completion.Length + 1;
        return true;
    }
    private bool HandleFormulaAssistKey(TextBox editor, Key key, ModifierKeys modifiers)
    {
        if (!IsFormulaDraft(editor) || modifiers != ModifierKeys.None) return false;
        if (_referenceGesture is not null) EndReferenceGesture();
        if (key == Key.F4)
        {
            if (FormulaEditing.ToggleReference(editor.Text, editor.SelectionStart, editor.SelectionLength) is { } change) ReplaceDraft(editor, change, true);
            return true;
        }
        if (_formulaHelp.IsOpen && _assistEditor == editor)
        {
            if (key == Key.Escape) { CloseFormulaHelp(); return true; }
            if (_completion is not null && key is Key.Up or Key.Down)
            { _formulaSuggestions.SelectedIndex = (_formulaSuggestions.SelectedIndex + (key == Key.Down ? 1 : _completion.Names.Length - 1)) % _completion.Names.Length; return true; }
            if (key == Key.Tab && AcceptCompletion()) return true;
        }
        return false;
    }

    private bool BeginReferenceGesture(CellPosition position)
    {
        if (Keyboard.FocusedElement is not TextBox editor || editor != FormulaBar && editor != _grid.Editor || !IsFormulaDraft(editor)) return false;
        if (FormulaEditing.ReferenceSlot(editor.Text, editor.SelectionStart, editor.SelectionLength) is not { } slot) return false;
        CloseFormulaHelp(); _referenceGesture = new(editor, editor.Text, slot, position);
        UpdateReferenceGesture(position); return true;
    }
    private void UpdateReferenceGesture(CellPosition position)
    {
        if (_referenceGesture is not { } gesture) return;
        var range = GridRange.Between(gesture.Start, position);
        var text = FormulaEditing.Address(range);
        _suppressFormulaHelp = true;
        try
        {
            gesture.Editor.Text = gesture.Original.Remove(gesture.Slot.Start, gesture.Slot.Length).Insert(gesture.Slot.Start, text);
            gesture.Editor.Select(gesture.Slot.Start, text.Length);
        }
        finally { _suppressFormulaHelp = false; }
        StatusText.Text = $"引用 {text}；Enter 提交到原单元格，Esc 取消草稿。";
    }
    private void EndReferenceGesture()
    {
        if (_referenceGesture is not { } gesture) return;
        _referenceGesture = null;
        if (_grid.Control.IsMouseCaptured) _grid.Control.ReleaseMouseCapture();
        gesture.Editor.Focus(); UpdateFormulaHelp(gesture.Editor);
    }
    private void AutoSum_Click(object sender, RoutedEventArgs e) => BeginAutoSum();
    private void BeginAutoSum()
    {
        if (_busy || !CanCommitComposition() || !CommitEditors() || !GridSurface.BeginEdit()) return;
        var source = EntryNavigation.AutoSumSource(_document.Snapshot.Sheets[_sheetIndex], new(_selectedRow, _selectedColumn));
        var reference = source is { } range ? FormulaEditing.Address(range) : "";
        _grid.Editor.Text = "=SUM(" + reference + ")";
        _grid.Editor.Select(5, reference.Length); _grid.Editor.Focus();
        StatusText.Text = "已建议求和范围（SUM 包含隐藏行）；可点选修改，Enter 提交，Esc 取消。";
    }
}
