using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly Popup _formulaHelp = new() { Placement = PlacementMode.Bottom, IsLightDismissEnabled = false };
    private readonly ListBox _formulaSuggestions = new() { MaxHeight = 190, MinWidth = 280, Focusable = false };
    private readonly TextBlock _formulaHint = new() { Margin = new(6), MaxWidth = 600, TextWrapping = TextWrapping.Wrap };
    private FormulaCompletion? _completion;
    private CompositionTextBox? _assistEditor;
    private CompositionTextBox? _referenceEditor;
    private bool _changingDraft;
    private sealed record FormulaGesture(CompositionTextBox Editor, string Original, FormulaTextEdit Slot);
    private FormulaGesture? _formulaGesture;
    private bool IsFormulaDraft(CompositionTextBox editor) => !_workspace.IsBusy && !editor.IsComposing &&
        editor.Text?.StartsWith('=') == true && (editor == _formula ? _formulaDirty : _grid.IsEditing);

    private void InstallFormulaEditing()
    {
        var panel = new StackPanel(); panel.Children.Add(_formulaSuggestions); panel.Children.Add(_formulaHint);
        _formulaHelp.Child = new Border { Background = Brushes.White, BorderBrush = Brushes.Teal, BorderThickness = new(1), Child = panel };
        _book.Children.Insert(0, _formulaHelp);
        foreach (var editor in new[] { _formula, _grid.CellEditor })
        {
            editor.PropertyChanged += (_, e) =>
            {
                if (e.Property == TextBox.TextProperty || e.Property == TextBox.CaretIndexProperty || e.Property == TextBox.SelectionStartProperty)
                    UpdateFormulaHelp(editor);
            };
            editor.GotFocus += (_, _) => { _referenceEditor = editor; UpdateFormulaHelp(editor); };
            editor.LostFocus += (_, _) => { if (_formulaGesture is null) CloseFormulaHelp(); };
            editor.CompositionChanged += () => UpdateFormulaHelp(editor);
        }
        _formulaSuggestions.PointerPressed += (_, e) =>
        {
            if (_completion is not null && _formulaSuggestions.SelectedItem is string) { AcceptCompletion(); e.Handled = true; }
        };
        // Let the window see the full Alt chord before consuming Enter, or its
        // access-key handler will focus the menu when Alt is released.
        _book.AddHandler(KeyDownEvent, (_, e) =>
        {
            if (e.Handled || _grid.IsComposing || _formula.IsComposing) return;
            var editor = _formula.IsKeyboardFocusWithin ? _formula : _grid.CellEditor.IsKeyboardFocusWithin ? _grid.CellEditor : null;
            if (editor is null) return;
            if (e.Key == Key.Enter && e.KeyModifiers == KeyModifiers.Alt)
            {
                e.Handled = true;
                if (editor.IsReadOnly) return;
                if (editor == _grid.CellEditor && !_grid.IsEditing)
                {
                    if (!_grid.BeginEdit()) return;
                    editor.SelectionStart = editor.SelectionEnd = editor.Text?.Length ?? 0;
                }
                editor.SelectedText = editor.NewLine;
                return;
            }
            if (e.Key == Key.OemPlus && e.KeyModifiers == KeyModifiers.Alt) { BeginAutoSum(); e.Handled = true; return; }
            if (!IsFormulaDraft(editor) || e.KeyModifiers != KeyModifiers.None) return;
            if (e.Key == Key.F4)
            {
                var start = Math.Min(editor.SelectionStart, editor.SelectionEnd);
                if (FormulaEditing.ToggleReference(editor.Text!, start, Math.Abs(editor.SelectionEnd - editor.SelectionStart)) is { } change)
                    ReplaceDraft(editor, change, true);
                e.Handled = true;
            }
            else if (_formulaHelp.IsOpen && _assistEditor == editor)
            {
                if (e.Key == Key.Escape) { CloseFormulaHelp(); e.Handled = true; }
                else if (_completion is not null && e.Key is Key.Up or Key.Down)
                { _formulaSuggestions.SelectedIndex = (_formulaSuggestions.SelectedIndex + (e.Key == Key.Down ? 1 : _completion.Names.Length - 1)) % _completion.Names.Length; e.Handled = true; }
                else if (e.Key == Key.Tab) e.Handled = AcceptCompletion();
            }
        }, RoutingStrategies.Tunnel);
        _grid.ReferenceRequested = PickFormulaReference;
        _grid.ReferenceGestureEnded += () =>
        {
            if (_formulaGesture is not { } gesture) return;
            _formulaGesture = null; gesture.Editor.Focus(); UpdateFormulaHelp(gesture.Editor);
        };
        Deactivated += (_, _) => CloseFormulaHelp(); Closed += (_, _) => _formulaHelp.IsOpen = false;
    }
    private void CloseFormulaHelp() { _formulaHelp.IsOpen = false; _completion = null; _assistEditor = null; }
    private void UpdateFormulaHelp(CompositionTextBox editor)
    {
        if (_changingDraft || _formulaGesture is not null || _synchronizing) return;
        if (!editor.IsKeyboardFocusWithin || !IsFormulaDraft(editor)) { if (_assistEditor == editor) CloseFormulaHelp(); return; }
        _assistEditor = editor;
        var reference = FormulaEditing.ReferenceAt(editor.Text!, editor.CaretIndex);
        _grid.HighlightReference(reference?.Sheet is null || reference.Sheet.Equals(Document.Snapshot.Sheets[SheetIndex].Name, StringComparison.OrdinalIgnoreCase) ? reference?.Range : null);
        _completion = editor.SelectionStart == editor.SelectionEnd ? FormulaEditing.Complete(editor.Text!, editor.CaretIndex) : null;
        _formulaSuggestions.ItemsSource = _completion?.Names; _formulaSuggestions.SelectedIndex = _completion is null ? -1 : 0;
        _formulaSuggestions.IsVisible = _completion is not null;
        var hint = FormulaEditing.Hint(editor.Text!, editor.CaretIndex);
        _formulaHint.Text = hint is null ? UiText.Get("↑↓ 选择 · Tab 补全 · F4 切换引用") : UiText.Format($"{hint.Signature}\n当前第 {hint.Argument} 个参数 · F4 切换引用");
        _formulaHelp.PlacementTarget = editor; _formulaHelp.IsOpen = _completion is not null || hint is not null;
    }
    private void ReplaceDraft(CompositionTextBox editor, FormulaTextEdit change, bool select)
    {
        _changingDraft = true;
        try
        {
            editor.Text = editor.Text!.Remove(change.Start, change.Length).Insert(change.Start, change.Text);
            editor.CaretIndex = change.Start + change.Text.Length;
            editor.SelectionStart = select ? change.Start : editor.CaretIndex; editor.SelectionEnd = editor.CaretIndex;
        }
        finally { _changingDraft = false; }
        editor.Focus(); UpdateFormulaHelp(editor);
    }
    private bool AcceptCompletion()
    {
        if (_assistEditor is not { } editor || _completion is not { } completion || _formulaSuggestions.SelectedItem is not string name || !IsFormulaDraft(editor)) return false;
        var after = completion.Start + completion.Length;
        while (after < editor.Text!.Length && char.IsWhiteSpace(editor.Text[after])) after++;
        var suffix = after < editor.Text!.Length && editor.Text[after] == '(' ? "" : "(";
        ReplaceDraft(editor, new(completion.Start, completion.Length, name + suffix), false);
        if (suffix.Length == 0) editor.CaretIndex = after + name.Length - completion.Length + 1;
        return true;
    }
    private bool PickFormulaReference(GridRange range)
    {
        if (_formulaGesture is null)
        {
            // The focusable grid receives focus before its pointer handler runs. Keep
            // the last text editor so a formula-bar click still edits that draft.
            if (_referenceEditor is not { } editor || !IsFormulaDraft(editor)) return false;
            var start = Math.Min(editor.SelectionStart, editor.SelectionEnd);
            if (FormulaEditing.ReferenceSlot(editor.Text!, start, Math.Abs(editor.SelectionStart - editor.SelectionEnd)) is not { } slot) return false;
            _formulaGesture = new(editor, editor.Text!, slot); CloseFormulaHelp();
        }
        var gesture = _formulaGesture; var text = FormulaEditing.Address(range);
        _changingDraft = true;
        try
        {
            gesture.Editor.Text = gesture.Original.Remove(gesture.Slot.Start, gesture.Slot.Length).Insert(gesture.Slot.Start, text);
            gesture.Editor.CaretIndex = gesture.Slot.Start + text.Length;
            gesture.Editor.SelectionStart = gesture.Slot.Start; gesture.Editor.SelectionEnd = gesture.Editor.CaretIndex;
        }
        finally { _changingDraft = false; }
        _grid.HighlightReference(range); return true;
    }
    private void BeginAutoSum()
    {
        if (!CommitEditors() || !_grid.BeginEdit()) return;
        var range = EntryNavigation.AutoSumSource(Document.Snapshot.Sheets[SheetIndex], _grid.Current);
        var reference = range is { } source ? FormulaEditing.Address(source) : "";
        _grid.CellEditor.Text = "=SUM(" + reference + ")";
        _grid.CellEditor.SelectionStart = 5; _grid.CellEditor.SelectionEnd = 5 + reference.Length; _grid.CellEditor.Focus();
    }
}
