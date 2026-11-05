using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunSpecialSelectionSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            _workspace.New(); ShowBook();
            _workspace.Editor.SetCell(0, 0, 0, "1"); _workspace.Editor.SetCell(0, 2, 0, "2"); Synchronize();
            _grid.SelectRange(new(0, 0, 3, 0), new(0, 0));
            Check(((MenuItem)_menu.Items[1]!).Items.OfType<MenuItem>().Any(i => i.Header as string == UiText.Get("定位条件…") && i.IsEnabled), "Edit menu exposes Go To Special");
            var before = Document.Snapshot; var revision = before.Revision;
            _dialogDriver = w => { w.Close(false); return Task.CompletedTask; };
            await GoToSpecialAsync(); Check(_grid.SelectionRanges.Single().Bottom == 3 && Document.Snapshot.Revision == revision, "Cancel preserves selection and revision");
            async Task ChooseConstants(Window w)
            {
                _dialogDriver = ChooseConstants;
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                if (w.Title == UiText.Get("定位条件")) w.GetVisualDescendants().OfType<ComboBox>().First().SelectedIndex = 1;
                w.Close(true);
            }
            _dialogDriver = ChooseConstants; await GoToSpecialAsync();
            Check(_grid.SelectionRanges.Count == 2 && Document.Snapshot.Revision == revision, "Constant matching selects disjoint ranges without editing history");
            var stats = SelectionStatistics.Compute(Document.Snapshot, Document.Snapshot.ActiveSheetId, _grid.SelectionRanges.Select(Rectangle));
            Check(stats.Count == 2 && stats.Sum == 3, "Selection statistics consume the result");
            await ApplyStyleAsync(new() { Bold = true });
            Check(new[] { 0, 2 }.All(row => Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(row, 0)].Bold == true), "Formatting applies to each matched region");
            await History(false);
            await ClearAsync(); Check(Document.GetCellInput(0, 0, 0) == "" && Document.GetCellInput(0, 2, 0) == "", "Clear applies to every matched region");
            await History(false); Check(Document.GetCellInput(0, 0, 0) == "1" && Document.GetCellInput(0, 2, 0) == "2", "One undo restores the entire clear");
            _grid.SelectRange(new(0, 0, 3, 0), new(0, 0));
            async Task Stale(Window w)
            {
                _dialogDriver = Stale;
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                if (w.Title == UiText.Get("定位条件")) w.GetVisualDescendants().OfType<ComboBox>().First().SelectedIndex = 1;
                if (w.Title == UiText.Get("确认定位")) _workspace.Editor.SetCell(0, 5, 0, "newer");
                w.Close(true);
            }
            _dialogDriver = Stale; var rejected = false;
            try { await GoToSpecialAsync(); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.GetCellInput(0, 5, 0) == "newer" && _grid.SelectionRanges.Single().Bottom == 3, "Stale preview preserves newer edits and original selection");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
