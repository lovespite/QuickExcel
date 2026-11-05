using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunWindowClosingSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>(); Exception? error = null;
        try
        {
            await CheckCloseAfterFailedOpenAsync(directory, (condition, description) =>
            {
                if (!condition) throw new InvalidOperationException(description);
                checks.Add(description);
            });
        }
        catch (Exception ex) { error = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new
        { passed = error is null, checks, error = error?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true;
        (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }

    private async Task CheckCloseAfterFailedOpenAsync(string directory, Action<bool, string> check)
    {
        var occupied = Path.Combine(directory, "occupied-on-close.xlsq");
        using (var blank = WorkbookDocument.CreateBlank(_services))
            new TypedWorkbookStorage().Save(occupied, blank.Snapshot);
        using var fileLock = new FileStream(occupied, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var smokeMode = Program.SmokeMode;
        Program.SmokeMode = "manual";
        try
        {
            foreach (var scenario in new[] { "window", "exit-menu", "pending-open", "unsaved" })
            {
                var window = new MainWindow(Path.Combine(directory, "closing", scenario));
                var closed = false;
                window.Closed += (_, _) => closed = true;
                try
                {
                    window.Show();
                    await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                    await window._operation;
                    var files = new SmokeFiles { NextOpen = occupied };
                    window._files = files;
                    var original = window.Document;
                    Task? pending = null;
                    var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                    if (scenario == "pending-open")
                    {
                        pending = window.PerformAsync(async () => { await release.Task; await window.OpenAsync(); });
                        window.Close();
                        check(!closed && window._closePending, "Close waits for the pending file operation");
                        release.SetResult();
                        await pending;
                    }
                    else
                    {
                        await window.PerformAsync(window.OpenAsync);
                        check(ReferenceEquals(original, window.Document) && !window._workspace.IsBusy && !window._working &&
                            window._workspace.FilePath is null && !string.IsNullOrEmpty(window._status.Text),
                            $"{scenario}: occupied file reports an error and preserves the current document");
                        if (scenario == "unsaved")
                        {
                            window.ShowBook();
                            window._grid.BeginEdit(); window._grid.CellEditor.Text = "keep this draft";
                            var prompts = 0;
                            window._discardChoice = () => { prompts++; return Task.FromResult(SaveDecision.Cancel); };
                            window.Close();
                            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                            check(!closed && prompts == 1 && window.Document.IsDirty && !window._grid.IsEditing &&
                                window.Document.GetCellInput(0, 0, 0) == "keep this draft",
                                "Failed open does not bypass draft commit or the unsaved cancel choice");
                            window._discardChoice = () => Task.FromResult(SaveDecision.Save);
                            window.Close();
                            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                            check(!closed && window.Document.IsDirty, "Cancelling Save As still cancels window close after failed open");
                            files.NextSave = Path.Combine(directory, "saved-on-close.xlsq");
                        }
                        if (scenario == "exit-menu")
                            ((MenuItem)window._menu.Items[0]!).Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get("退出")))
                                .RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                        else window.Close();
                    }
                    for (var attempt = 0; !closed && attempt < 200; attempt++) await Task.Delay(10);
                    check(closed, $"{scenario}: a single close request exits after the occupied-file error");
                    if (scenario == "unsaved")
                    {
                        using var saved = new WorkbookDocument(new TypedWorkbookStorage().Open(files.NextSave!), _services);
                        check(saved.GetCellInput(0, 0, 0) == "keep this draft", "Close saves pending edits and releases the saved file lease");
                    }
                }
                finally { if (!closed) { window._allowClose = true; window.Close(); } }
            }
        }
        finally { Program.SmokeMode = smokeMode; }
    }
}
