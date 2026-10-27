using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCloseSmokeTestsAsync(string folder)
    {
        var windows = new List<MainWindow>();
        MainWindow CreateWindow(WorkbookServices? services = null, string? path = null)
        {
            var window = new MainWindow
            {
                Owner = this, ShowInTaskbar = false, ShowActivated = false, Opacity = 0,
                _smokeMode = true
            };
            windows.Add(window);
            window.InstallDocument(WorkbookDocument.CreateBlank(services ?? _services), path, "关闭回归.xlsq");
            window.Show();
            return window;
        }

        static async Task CloseAndWaitAsync(MainWindow window, bool repeat = false)
        {
            var closed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var count = 0;
            window.Closed += (_, _) => { count++; closed.TrySetResult(); };
            window.Close(); // Exercise WPF Closing; do not set _closingApproved here.
            if (repeat) window.Close();
            await closed.Task.WaitAsync(TimeSpan.FromSeconds(5));
            if (count != 1 || !window._windowClosed || !window._document.Session.IsClosed)
                throw new Exception("Actual window close did not dispose exactly once.");
        }

        static async Task WaitForDecisionAsync(MainWindow window)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            do
            {
                await window.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                if (!window._closeRequestPending && !window._busy) return;
                await Task.Delay(10);
            } while (DateTime.UtcNow < deadline);
            throw new Exception("Close confirmation did not finish.");
        }

        try
        {
            var clean = CreateWindow();
            await CloseAndWaitAsync(clean, repeat: true);

            var blockedCalculator = new SmokeFailingCalculator();
            var discarded = CreateWindow(new(_services.Storage, blockedCalculator));
            discarded.ApplyValues([(0, 0, "1")]);
            discarded._smokeDiscardDecision = MessageBoxResult.No;
            discarded.QueueRecalculation();
            await CloseAndWaitAsync(discarded, repeat: true);
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            if (blockedCalculator.Attempts != 0)
                throw new Exception("Queued automatic calculation ran during or after discard-close.");

            var canceled = CreateWindow();
            canceled.ApplyValues([(0, 0, "2")]);
            canceled._smokeDiscardDecision = MessageBoxResult.Cancel;
            canceled.Close();
            await WaitForDecisionAsync(canceled);
            if (!canceled.IsVisible || canceled._document.Session.IsClosed || !canceled._document.IsDirty)
                throw new Exception("Cancel close lost the editable document.");
            canceled.ApplyValues([(0, 0, "3")]);
            canceled._smokeDiscardDecision = MessageBoxResult.No;
            await CloseAndWaitAsync(canceled);

            var invalid = CreateWindow();
            var beforeInvalid = invalid._document.Snapshot.Revision;
            invalid.FormulaBar.Text = new string('x', WorkbookLimits.MaxCellTextLength + 1);
            invalid.Close();
            await WaitForDecisionAsync(invalid);
            if (!invalid.IsVisible || !invalid._formulaPending || invalid._document.Snapshot.Revision != beforeInvalid)
                throw new Exception("Invalid pending input was lost during close.");
            invalid.FormulaBar.Text = "5";
            invalid._smokeDiscardDecision = MessageBoxResult.No;
            await CloseAndWaitAsync(invalid);

            var savePath = Path.Combine(folder, "close-saved.xlsq");
            var saved = CreateWindow(path: savePath);
            saved.ApplyValues([(0, 0, "40"), (0, 1, "=A1*2")]);
            saved._smokeDiscardDecision = MessageBoxResult.Yes;
            await CloseAndWaitAsync(saved);
            using (var reopened = WorkbookDocument.Open(savePath, _services))
                if (reopened.GetCellValue(0, 0, 0) != "40" || reopened.GetCellValue(0, 0, 1) != "80")
                    throw new Exception("Save-on-close lost committed inputs or formula caches.");

            var failed = CreateWindow(new(new SmokeFailingStorage(_services.Storage), _services.Calculator), savePath);
            failed.ApplyValues([(0, 0, "99")]);
            failed._smokeDiscardDecision = MessageBoxResult.Yes;
            failed.Close();
            await WaitForDecisionAsync(failed);
            if (!failed.IsVisible || failed._document.Session.IsClosed || !failed._document.IsDirty)
                throw new Exception("Failed save-on-close disposed or marked the document saved.");
            using (var previous = WorkbookDocument.Open(savePath, _services))
                if (previous.GetCellValue(0, 0, 0) != "40") throw new Exception("Failed close-save changed the previous file.");
            failed.ApplyValues([(0, 0, "100")]);
            failed._smokeDiscardDecision = MessageBoxResult.No;
            await CloseAndWaitAsync(failed);

            var busy = CreateWindow();
            var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var work = busy.WorkAsync("关闭期间忙碌回归…", () =>
            {
                started.SetResult();
                release.Task.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();
            });
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            busy.Close();
            var refusedWhileBusy = busy.IsVisible && !busy._closeRequestPending && !busy._document.Session.IsClosed;
            release.SetResult();
            if (!await work || !refusedWhileBusy) throw new Exception("A busy window was closed while work was running.");
            await CloseAndWaitAsync(busy);

            var vetoed = CreateWindow();
            CancelEventHandler veto = (_, args) => args.Cancel = true;
            vetoed.Closing += veto;
            vetoed.Close();
            await WaitForDecisionAsync(vetoed);
            if (!vetoed.IsVisible || vetoed._document.Session.IsClosed || vetoed._closingApproved)
                throw new Exception("A canceled final close disposed the session or retained approval.");
            vetoed.Closing -= veto;
            await CloseAndWaitAsync(vetoed);

            File.WriteAllText(Path.Combine(folder, "close-smoke-result.txt"),
                "PASS: actual WPF Close on clean document, discard with queued calculation, cancel then edit/retry, invalid pending input, save then close, failed save retains document/file, busy close refused, final-close veto retains session; duplicate requests do not reenter.");
        }
        finally
        {
            foreach (var window in windows.Where(window => !window._windowClosed))
            {
                window._closingApproved = true; // Cleanup only; success assertions use the public close path.
                window.Close();
            }
        }
    }

    private sealed class SmokeFailingStorage(IWorkbookStorage inner) : IWorkbookStorage
    {
        public WorkbookSnapshot Open(string path) => inner.Open(path);
        public void Save(string path, WorkbookSnapshot snapshot) => throw new IOException("Smoke: simulated save failure.");
        public void Export(string path, WorkbookSnapshot snapshot) => inner.Export(path, snapshot);
    }
}
