using System.Diagnostics;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunRecoverySmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            _workspace.New(); ShowBook(); _workspace.Editor.SetCell(0, 0, 0, "Recover 中文"); Synchronize();
            await CheckStatusMessageLayoutAsync(Check);
            _recoveryTimer.Interval = TimeSpan.FromMilliseconds(50); StartRecovery();
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (_lastRecovery is null && DateTime.UtcNow < deadline) await Task.Delay(20);
            _recoveryEnabled = false; _recoveryTimer.Stop(); await _recoveryWrite;
            Check(_lastRecovery?.Revision == Document.Snapshot.Revision, "Recovery timer records current session revision");
            Check((await Task.Run(_recoveryStore.Scan)).Items.Count == 0, "Running instance recovery lease is excluded");
            _recoveryLease!.Dispose(); _recoveryLease = null; _recoverySession = Guid.Empty;
            _workspace.New(); ShowHome();
            _dialogDriver = w => { w.Close(true); return Task.CompletedTask; }; await ShowRecoveryAsync(true);
            Check(Document.GetCellValue(0, 0, 0) == "Recover 中文" && Document.IsDirty && _workspace.FilePath is null && !Document.ScriptsEnabled,
                "Startup recovery creates unsaved detached document without enabling scripts");
            _files = new RecoveryDialogs(Path.Combine(directory, "saved.xlsq"));
            Check(await SaveAsync(false), "Recovered workbook requires and accepts save as");
            Check((await Task.Run(_recoveryStore.Scan)).Items.Count == 0 && _recoveredFrom is null && _recoveryLease is null, "Successful save clears both recovered and current drafts");
            _workspace.Editor.SetCell(0, 1, 0, "discard"); await CaptureRecoveryAsync();
            _discardChoice = () => Task.FromResult(SaveDecision.Discard); Check(await ConfirmDiscardAsync(), "Explicit discard closes recovery lifecycle");
            Check(_recoveryLease is null && (await Task.Run(_recoveryStore.Scan)).Items.Count == 0, "Discard removes current draft generations");
            await ClearRecoveryAsync();
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; _discardChoice = null; await ClearRecoveryAsync(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
    private async Task CheckStatusMessageLayoutAsync(Action<bool, string> check)
    {
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var status = _status.Text;
        var recovery = _recoveryStatus.Text;
        var gridHeight = _grid.Bounds.Height;
        var footer = (DockPanel)_status.Parent!;
        var footerHeight = footer.Bounds.Height;
        var message = "恢复草稿写入失败：此操作会丢失尚未支持的内容。\n" +
            string.Join("\n", Enumerable.Range(0, 24).Select(i => $"xl/worksheets/sheet1.xml: 未建模元素 {i}；请查看兼容性报告。"));
        try
        {
            _status.Text = message;
            _recoveryStatus.Text = message;
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            check(footer.Bounds.Height <= footerHeight + 1 && _grid.Bounds.Height >= gridHeight - 1,
                "Multiline errors keep the status bar compact and preserve worksheet height");
            check(_status.Text == message && _recoveryStatus.Text == message &&
                Equals(ToolTip.GetTip(_status), message) && Equals(ToolTip.GetTip(_recoveryStatus), message),
                "Truncated status messages retain every diagnostic line in their tooltips");
            _status.Text = "就绪"; _recoveryStatus.Text = "恢复草稿：12:00:00";
            check(Equals(ToolTip.GetTip(_status), _status.Text) && Equals(ToolTip.GetTip(_recoveryStatus), _recoveryStatus.Text),
                "Status tooltips update after an error clears");
        }
        finally
        {
            _status.Text = status; _recoveryStatus.Text = recovery;
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        }
    }
    private async Task RunRecoveryBenchmarkAsync(string directory)
    {
        Directory.CreateDirectory(directory); Exception? failure = null; var samples = new List<double>(); var total = new Stopwatch();
        try
        {
            // Separate measurement run; no functional assertions or timing gates mixed into the smoke tests.
            _workspace.New(); ShowBook();
            var sheet = Document.Snapshot.ActiveSheetId;
            Document.Session.Execute(new SetCellsCommand(Enumerable.Range(0, 20000).Select(i => new CellEdit(new(sheet, i, 0), CellContent.FromLiteral(CellValue.FromText("记录 " + i))))));
            Synchronize(); total.Start(); var capture = CaptureRecoveryAsync();
            while (!capture.IsCompleted)
            {
                var latency = Stopwatch.StartNew(); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); samples.Add(latency.Elapsed.TotalMilliseconds);
                await Task.Delay(10);
            }
            await capture; total.Stop();
            if (_lastRecovery?.Revision != Document.Snapshot.Revision) throw new InvalidOperationException(_recoveryStatus.Text);
        }
        catch (Exception ex) { failure = ex; }
        finally { await ClearRecoveryAsync(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, error = failure?.ToString(), cells = 20000, elapsedMs = total.Elapsed.TotalMilliseconds, dispatcherSamplesMs = samples, maxDispatcherMs = samples.DefaultIfEmpty().Max(), targetWindowMs = 60000, withinTarget = total.Elapsed.TotalSeconds < 60 }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
    private sealed class RecoveryDialogs(string destination) : ILocalFileDialogs
    {
        public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request) => Task.FromResult<LocalFileSelection?>(null);
        public Task<LocalFileSelection?> SaveAsync(FileDialogRequest request) => Task.FromResult<LocalFileSelection?>(new(destination));
    }
}
