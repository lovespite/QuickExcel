using Avalonia.Controls;
using Avalonia.Input;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Func<Task>? ExcelShortcut(KeyEventArgs e)
    {
        if (e.KeyModifiers == KeyModifiers.None)
            return e.Key switch { Key.F12 => async () => await SaveAsync(true), Key.F5 => NavigateAsync, _ => null };
        if (e.KeyModifiers == KeyModifiers.Shift)
            return e.Key switch { Key.F12 => async () => await SaveAsync(false), Key.F11 => AddSheetAsync, _ => null };
        if (e.KeyModifiers is not (KeyModifiers.Control or KeyModifiers.Meta)) return null;
        if (e.Key == Key.P) return OpenPrintPreviewAsync;
        if (e.Key == Key.F12) return OpenAsync;
        if (e.Key == Key.F4) return async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } };
        // Native text editing keeps ownership of its shortcuts; formatting applies to cell selections.
        if (_grid.IsEditing || _formula.IsKeyboardFocusWithin || e.Source is TextBox) return null;
        return e.Key switch
        {
            Key.B => () => ApplyStyleAsync(new(Bold: !CurrentStyle.Bold)),
            Key.I => () => ApplyStyleAsync(new(Italic: !CurrentStyle.Italic)),
            Key.U => () => ApplyStyleAsync(new(Underline: CurrentStyle.Underline is null or "none" ? "single" : "none")),
            Key.D5 or Key.NumPad5 => () => ApplyStyleAsync(new(Strike: !CurrentStyle.Strike)),
            _ => null
        };
    }
}
