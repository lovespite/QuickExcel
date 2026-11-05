using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckRecentFileFeaturesAsync(string directory, Action<bool, string> check, Func<Control, string, Task> capture)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        static void Choose(Window window, string label) => Click(window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, UiText.Get(label))));
        async Task Idle()
        {
            for (var attempt = 0; _working && attempt < 1000; attempt++) await Task.Delay(10);
            if (_working) throw new TimeoutException("History operation did not release the input gate.");
            await Layout();
        }
        async Task Open(string path)
        {
            await Idle(); Click(RecentOpenButtons().Single(b => Equals(b.Tag, path))); await _operation; await Idle();
        }
        async Task Pin(string path)
        {
            var button = _recent.GetVisualDescendants().OfType<ToggleButton>().Single(b => Equals(b.Tag, path));
            check(button.Content is global::Avalonia.Controls.Shapes.Path && ToolTip.GetTip(button) is string &&
                AutomationProperties.GetName(button)?.Contains(Path.GetFileName(path)) == true, "Pin is an icon with a tooltip and accessible file name");
            await Idle(); Click(button); await _operation; await Idle();
        }
        var config = Path.Combine(directory, "settings", "recent-files.json");
        var validPath = Path.Combine(directory, "年度预算与项目计划.xlsq");
        var missing = _recentFiles.Paths.Single(p => p.Contains("HistoryFolder-4", StringComparison.Ordinal));
        var missingDirectory = _recentFiles.Paths.Single(p => p.Contains("HistoryFolder-5", StringComparison.Ordinal));
        Directory.CreateDirectory(Path.GetDirectoryName(missing)!);
        var document = Document;
        document.SetCell(0, 0, 0, "keep this unsaved workbook");
        var revision = document.Snapshot.Revision;
        var originalDiscard = _discardChoice; var unsavedPrompts = 0;
        _discardChoice = () => { unsavedPrompts++; return Task.FromResult(SaveDecision.Cancel); };
        try
        {
            await Pin(missing);
            check(_recentFiles.Entries[0] is { IsPinned: true } entry && entry.Path == missing && ReferenceEquals(Document, document) && unsavedPrompts == 0,
                "Pin moves the entry to the top without opening it or prompting about unsaved changes");
            RefreshRecentFiles(); await Layout();
            check(_recent.GetVisualDescendants().OfType<ToggleButton>().First().IsChecked == true, "Pin state survives a history reload");
            await Pin(missing);
            check(!_recentFiles.Entries.Single(e => e.Path == missing).IsPinned, "The icon can unpin an entry");
            await Pin(missing); await capture(this, "home-pinned-history");

            _recentSearch.Text = "historyfolder-4"; await Layout();
            var before = await File.ReadAllTextAsync(config);
            foreach (var close in new[] { false, true })
            {
                _dialogDriver = async window =>
                {
                    await Layout();
                    check(window.Title == UiText.Get("文件不存在") && window.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Text?.Contains(missing) == true),
                        "Missing-file dialog names the missing workbook");
                    if (close) window.Close(); else { await capture(window, "missing-history-file"); Choose(window, "保留记录"); }
                };
                await Open(missing);
                check(await File.ReadAllTextAsync(config) == before && RecentOpenButtons().Count() == 1 && unsavedPrompts == 0,
                    "Keeping or closing the missing-file prompt preserves history, search and unsaved work");
            }
            _dialogDriver = async window => { await Layout(); Choose(window, "移除记录"); };
            await Open(missing);
            var reopened = new RecentFileStore(config); reopened.Load();
            check(!reopened.Paths.Contains(missing) && _recent.Children.Single() is TextBlock && _recentSearch.Text == "historyfolder-4",
                "Confirmed removal persists, removes a pinned entry and refreshes the current search");
            _recentSearch.Text = ""; await Layout();
            _dialogDriver = async window => { await Layout(); Choose(window, "移除记录"); };
            await Open(missingDirectory);
            check(!_recentFiles.Paths.Contains(missingDirectory) && unsavedPrompts == 0 && ReferenceEquals(Document, document) && Document.Snapshot.Revision == revision && Document.IsDirty,
                "A missing parent directory uses the same removal prompt without changing the active document");

            _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            using (var locked = new FileStream(validPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                _dialogDriver = window => { window.Close(); throw new InvalidOperationException("Locked workbook was incorrectly reported as missing."); };
                try { await Open(validPath); throw new InvalidOperationException("A locked workbook unexpectedly opened."); }
                catch (IOException) { await Layout(); }
                check(_dialogDriver is not null && ReferenceEquals(Document, document) && _recentFiles.Paths.Contains(validPath),
                    "A sharing violation retains its ordinary error path and keeps the history entry");
                _dialogDriver = null;
            }
            var racing = Path.Combine(directory, "disappears-after-confirmation.xlsq");
            new TypedWorkbookStorage().Save(racing, Document.Snapshot); _recentFiles.Remember(racing); RefreshRecentFiles(); await Layout();
            _discardChoice = () => { File.Delete(racing); return Task.FromResult(SaveDecision.Discard); };
            _dialogDriver = async window => { await Layout(); Choose(window, "移除记录"); };
            await Open(racing);
            check(!_recentFiles.Paths.Contains(racing) && ReferenceEquals(Document, document), "A file deleted during the unsaved-work prompt can still be removed from history");

            await Pin(validPath);
            _dialogDriver = async window =>
            {
                await Layout();
                var limit = window.GetVisualDescendants().OfType<NumericUpDown>().Single(n => n.Name == "RecentFileLimit");
                check(limit.Value == 100 && limit.Minimum == 1 && limit.Maximum == 1000, "History settings expose the default and supported limit range");
                limit.Value = 2;
                Click(window.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "ApplyRecentFileLimit")); await Layout();
                check(_recentFiles.Limit == 2 && _recentFiles.Entries.Count(e => !e.IsPinned) == 2 && _recentFiles.Entries.Single(e => e.IsPinned).Path == validPath,
                    "Applying a lower limit trims ordinary history while preserving pinned files");
                await capture(window, "settings-history-limit"); window.Close(true);
            };
            Click(_start.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "ManageRecentFiles")); await _operation; await Layout();
            reopened.Load();
            check(reopened.Limit == 2 && reopened.Paths.SequenceEqual(_recentFiles.Paths) && RecentOpenButtons().Count() == 3 && Equals(RecentOpenButtons().First().Tag, validPath),
                "The limit and pinned-first ordering persist and immediately refresh the home page");
            _dialogDriver = async window =>
            {
                await Layout();
                check(window.GetVisualDescendants().OfType<NumericUpDown>().Single(n => n.Name == "RecentFileLimit").Value == 2,
                    "Reopening Settings restores the saved history limit");
                window.Close(true);
            };
            Click(_start.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "ManageRecentFiles")); await _operation;
        }
        finally { _dialogDriver = null; _discardChoice = originalDiscard; _recentSearch.Text = ""; _workspace.New(); }
    }
}
