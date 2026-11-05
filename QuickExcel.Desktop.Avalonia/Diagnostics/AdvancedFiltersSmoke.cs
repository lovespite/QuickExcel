using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunAdvancedFiltersSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            _workspace.New(); ShowBook(); _workspace.Editor.SetCells(0, GridClipboard.Prepare("Amount\n10\n20\n30", new(0, 0, 0, 0), false)); Synchronize();
            var id = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellStyleCommand([new(id, 1, 0), new(id, 3, 0)], new(FillColor: "FFFF0000")));
            _grid.SelectRange(new(0, 0, 3, 0), new(0, 0)); await EstablishFilterAsync(); _grid.Select(new(0, 0));
            async Task ChooseColor(FilterPanel panel)
            {
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                var button = panel.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "FilterColor");
                button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
            _filterDriver = ChooseColor; _dialogDriver = w => { w.Close(false); return Task.CompletedTask; };
            var revision = Document.Snapshot.Revision; await FilterAsync(0);
            Check(Document.Snapshot.Revision == revision, "Cancel color filtering preserves the existing filter");
            async Task Red(Window w)
            {
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                ((TextBox)((FieldsDialog)w).Input("color")).Text = "#FF0000"; w.Close(true);
            }
            _filterDriver = ChooseColor; _dialogDriver = Red; await FilterAsync(0);
            Check(Document.Snapshot.Sheets[0].Data.Filter!.Columns[0].Color == new AutoFilterColor(false, "FFFF0000"), "Normal column filter menu applies a fill color criterion");
            Check(WorksheetFilterEngine.Evaluate(Document.Snapshot, id).SetEquals([2]), "Grid filter evaluation hides nonmatching base colors");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.Filter!.Columns.Count == 0, "One undo removes only the color criterion");
            _filterDriver = ChooseColor;
            _dialogDriver = async w =>
            {
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                ((TextBox)((FieldsDialog)w).Input("color")).Text = "#FF0000";
                _workspace.Editor.SetCell(0, 8, 0, "newer"); w.Close(true);
            };
            var rejected = false; try { await FilterAsync(0); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.Snapshot.Sheets[0].Data.Filter!.Columns.Count == 0 && Document.GetCellInput(0, 8, 0) == "newer", "Stale color dialog preserves newer edits without applying criteria");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; _filterDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
