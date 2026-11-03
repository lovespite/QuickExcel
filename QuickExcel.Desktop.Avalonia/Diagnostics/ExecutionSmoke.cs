using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunExecutionSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>(); Exception? failure = null; Window? secondary = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        async Task Visible()
        {
            for (var i = 0; i < 100 && !_execution.IsOverlayVisible; i++) await Task.Delay(10);
            await Layout();
            Check(_execution.IsOverlayVisible, "Work lasting over 100 ms displays the overlay");
        }
        void Restored() => Check(!_execution.IsRunning && !_execution.IsOverlayVisible && _execution.Children[0].IsHitTestVisible &&
            _book.IsEffectivelyEnabled && _ribbon.IsEffectivelyEnabled && _menu.IsEffectivelyEnabled,
            "Completion restores input without changing enabled styles or leaving an overlay");
        try
        {
            _eventTimer.Stop(); _workspace.New(); ShowBook(); await Layout();
            var footer = (DockPanel)_status.Parent!; var footerHeight = footer.Bounds.Height;
            Check(!footer.Children.OfType<Button>().Any(), "The status bar contains no stop button");

            var fastVisible = false;
            for (var i = 0; i < 5; i++)
                await ScriptWorkAsync(async _ =>
                {
                    Check(_book.IsEffectivelyEnabled && _ribbon.IsEffectivelyEnabled && _menu.IsEffectivelyEnabled && _automationLoad!.IsEnabled,
                        "Short code execution preserves the workbook and button appearance");
                    await Task.Yield(); fastVisible |= _execution.IsOverlayVisible;
                });
            await Task.Delay(140); Restored();
            Check(!fastVisible, "Consecutive fast operations never display a stale delayed overlay");

            _workspace.Editor.Execute(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 0, 0), CellInput.Parse("=1+2"))]));
            var enabledDuringCalculation = false;
            void OnBusy()
            {
                if (_workspace.IsBusy) enabledDuringCalculation = _execution.IsRunning && _book.IsEffectivelyEnabled && _ribbon.IsEffectivelyEnabled && _menu.IsEffectivelyEnabled;
            }
            _workspace.Changed += OnBusy;
            try { await CalculateAsync(); } finally { _workspace.Changed -= OnBusy; }
            Check(enabledDuringCalculation && Document.GetCellValue(0, 0, 0) == "3", "Real recalculation preserves enabled styles and publishes the calculated value");
            Restored();

            var gate = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var running = ScriptWorkAsync(token => gate.Task.WaitAsync(token));
            try
            {
                Check(!_execution.IsOverlayVisible, "The loading overlay is initially hidden");
                var duplicates = 0; await PerformAsync(() => { duplicates++; return Task.CompletedTask; });
                Check(duplicates == 0 && !await CommitCellAsync(0, new(0, 0), "99"), "The delay does not admit duplicate commands or concurrent cell commits");
                await Visible();
                Check(_execution.Elapsed.Text!.EndsWith(" ms") && _execution.Cancel.IsEffectivelyEnabled, "The overlay shows elapsed milliseconds and an enabled cancel button");
                var elapsed = _execution.Elapsed.Text; await Task.Delay(80);
                Check(_execution.Elapsed.Text != elapsed, "The elapsed timer continues updating during execution");
                Check(Math.Abs(footer.Bounds.Height - footerHeight) < .1, "Long execution does not resize the status bar");
                using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(Bounds.Width), (int)Math.Ceiling(Bounds.Height)), new Vector(96, 96));
                bitmap.Render(this); bitmap.Save(Path.Combine(directory, "execution-overlay.png"), PngBitmapEncoderOptions.Default);
            }
            finally { gate.TrySetResult(); await running; }
            Restored();

            running = ScriptWorkAsync(token => Task.Delay(Timeout.Infinite, token));
            RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Escape });
            try { await running; Check(false, "Escape must cancel"); } catch (OperationCanceledException) { }
            Check(!_execution.IsOverlayVisible, "Escape cancels even before the overlay delay expires"); Restored();

            running = ScriptWorkAsync(token => Task.Delay(Timeout.Infinite, token));
            await Visible(); _execution.Cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            try { await running; Check(false, "The cancel button must cancel"); } catch (OperationCanceledException) { }
            Restored();

            var fail = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            running = ScriptWorkAsync(_ => fail.Task); await Visible(); fail.SetException(new InvalidOperationException("expected execution failure"));
            try { await running; } catch (InvalidOperationException error) when (error.Message == "expected execution failure") { }
            Restored();

            // Exercise the second window with native-surface visibility hooks, without starting a browser or worker.
            secondary = new Window { Width = 600, Height = 400 };
            var input = new TextBox { Text = "preserved draft" }; var surface = new ContentControl();
            var content = new StackPanel(); content.Children.Add(input); content.Children.Add(surface);
            var composing = false;
            var scriptFeedback = _scriptExecution = new ExecutionOverlay(secondary, content, () => composing, visible => surface.IsVisible = !visible);
            secondary.Content = scriptFeedback; secondary.Show(this); await Layout(); input.Focus();
            running = ScriptWorkAsync(token => Task.Delay(Timeout.Infinite, token));
            Check(surface.IsVisible && !scriptFeedback.IsOverlayVisible && input.IsEffectivelyEnabled, "The script window stays visually unchanged before the delay");
            await Visible();
            for (var i = 0; i < 100 && !scriptFeedback.IsOverlayVisible; i++) await Task.Delay(10);
            Check(scriptFeedback.IsOverlayVisible && !surface.IsVisible, "The script overlay covers the native editor surface after the delay");
            composing = true; secondary.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Escape });
            Check(!_scriptCancellation!.IsCancellationRequested, "Composition Escape is not treated as execution cancellation");
            composing = false; secondary.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Escape });
            try { await running; Check(false, "Script-window Escape must cancel"); } catch (OperationCanceledException) { }
            Check(!scriptFeedback.IsOverlayVisible && surface.IsVisible && content.IsHitTestVisible && input.Text == "preserved draft", "Script cancellation restores the editor surface and its draft");
            Restored();
        }
        catch (Exception error) { failure = error; }
        finally { _scriptCancellation?.Cancel(); _scriptExecution = null; secondary?.Close(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
