using Avalonia.Controls;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
#if QUICKEXCEL_DIAGNOSTICS
    private Func<Window, Task>? _dialogDriver;
#endif
    internal async Task<T> ShowOwnedAsync<T>(Window window, Window? owner = null)
    {
        Exception? failure = null;
#if QUICKEXCEL_DIAGNOSTICS
        if (Program.SmokeDirectory is not null && _dialogDriver is { } driver)
        {
            _dialogDriver = null;
            window.Opened += async (_, _) =>
            {
                try { await driver(window); }
                catch (Exception ex) { failure = ex; window.Close(default(T)); }
            };
        }
#endif
        var result = await window.ShowDialog<T>(owner ?? this);
        if (failure is not null) throw new InvalidOperationException(UiText.Get("对话框专项检查失败。"), failure);
        return result;
    }
}
