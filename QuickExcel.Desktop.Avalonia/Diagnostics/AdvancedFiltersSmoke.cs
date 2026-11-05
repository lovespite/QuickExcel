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
            _filterDriver = async panel =>
            {
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                panel.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "FilterTop").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            _dialogDriver = w => { ((TextBox)((FieldsDialog)w).Input("count")).Text = "2"; w.Close(true); return Task.CompletedTask; };
            await FilterAsync(0);
            Check(Document.Snapshot.Sheets[0].Data.Filter!.Columns[0].Top == new AutoFilterTop(2), "Column filter menu applies top N criteria");
            Check(WorksheetFilterEngine.Evaluate(Document.Snapshot, id).SetEquals([1]), "Top N hides rows below the numeric threshold");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.Filter!.Columns.Count == 0, "Top N criteria are one undoable change");
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("46331\n46356\n46357", new(1, 0, 3, 0), false));
            await ExecuteAsync(new SetCellStyleCommand([new(id, 1, 0), new(id, 2, 0), new(id, 3, 0)], new(NumberFormatId: 0, NumberFormatCode: "yyyy-mm-dd")));
            async Task ChooseDates(FilterPanel panel)
            {
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                panel.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "FilterDateGroups").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
            _filterDriver = ChooseDates; _dialogDriver = w => { w.Close(false); return Task.CompletedTask; };
            revision = Document.Snapshot.Revision; await FilterAsync(0); Check(Document.Snapshot.Revision == revision, "Cancel date grouping preserves criteria");
            _filterDriver = ChooseDates; _dialogDriver = w => { ((TextBox)((FieldsDialog)w).Input("dates")).Text = "2026-11"; w.Close(true); return Task.CompletedTask; };
            await FilterAsync(0);
            Check(WorksheetFilterEngine.Evaluate(Document.Snapshot, id).SetEquals([3]), "Column menu filters a calendar month without matching December");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.Filter!.Columns.Count == 0, "Date group is one undoable edit");
            _filterDriver = ChooseDates; _dialogDriver = w =>
            {
                ((TextBox)((FieldsDialog)w).Input("dates")).Text = "2026";
                _workspace.Editor.SetCell(0, 8, 0, "date newer"); w.Close(true); return Task.CompletedTask;
            };
            rejected = false; try { await FilterAsync(0); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.Snapshot.Sheets[0].Data.Filter!.Columns.Count == 0, "Stale date dialog does not change filter criteria");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; _filterDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
