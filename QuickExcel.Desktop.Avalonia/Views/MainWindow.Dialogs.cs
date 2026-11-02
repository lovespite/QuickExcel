using Avalonia.Controls;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Func<Window, Task>? _dialogDriver;
    internal async Task<T> ShowOwnedAsync<T>(Window window)
    {
        Exception? failure = null;
        if (Program.SmokeDirectory is not null && _dialogDriver is { } driver)
        {
            _dialogDriver = null;
            window.Opened += async (_, _) =>
            {
                try { await driver(window); }
                catch (Exception ex) { failure = ex; window.Close(default(T)); }
            };
        }
        var result = await window.ShowDialog<T>(this);
        if (failure is not null) throw new InvalidOperationException("对话框专项检查失败。", failure);
        return result;
    }
}
