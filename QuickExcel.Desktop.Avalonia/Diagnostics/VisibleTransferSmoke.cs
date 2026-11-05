using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunVisibleTransferSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            var memory = new MemoryClipboard(); _clipboard = memory; _workspace.New(); ShowBook(); var id = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromText("00123"))),
                new(new(id, 2, 0), CellContent.FromLiteral(CellValue.FromNumber(2))), new(new(id, 5, 2), CellContent.FromLiteral(CellValue.FromText("keep")))]));
            await ExecuteAsync(new SetAxisMetadataCommand(id, SheetAxis.Row, 1, 1, hidden: true));
            await ExecuteAsync(new SetAxisMetadataCommand(id, SheetAxis.Row, 5, 1, hidden: true));
            _grid.SelectRange(new(0, 0, 2, 0), new(0, 0));
            Check(((MenuItem)_menu.Items[1]!).Items.OfType<MenuItem>().Any(i => i.Header as string == UiText.Get("复制可见单元格…") && i.IsEnabled), "Edit menu exposes visible copy");
            _dialogDriver = w => { w.Close(true); return Task.CompletedTask; }; await CopyVisibleAsync();
            Check(memory.Value.SpreadsheetXml is not null && _visibleCapture!.Data.Rows == 2, "Visible copy publishes compact typed rows and styles");
            _grid.SelectRange(new(4, 2, 6, 2), new(4, 2)); var revision = Document.Snapshot.Revision;
            _dialogDriver = w => { w.Close(false); return Task.CompletedTask; }; await PasteVisibleAsync();
            Check(Document.Snapshot.Revision == revision, "Cancel paste preserves all rows");
            _dialogDriver = w => { w.Close(true); return Task.CompletedTask; }; await PasteVisibleAsync();
            Check(Document.GetCellValue(0, 4, 2) == "00123" && Document.GetCellValue(0, 6, 2) == "2" && Document.GetCellValue(0, 5, 2) == "keep", "Paste maps visible rows and preserves the hidden row");
            await History(false); Check(Document.GetCellInput(0, 4, 2) == "" && Document.GetCellInput(0, 6, 2) == "" && Document.GetCellInput(0, 5, 2) == "keep", "One undo restores every visible destination");
            var rejected = false; try { await PasteVisibleAsync(); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected, "A changed source revision requires recopying");
            memory.Value = memory.Value with { Token = null }; // A different instance receives only portable representations.
            _dialogDriver = w => { _workspace.Editor.SetCell(0, 9, 0, "newer"); w.Close(true); return Task.CompletedTask; };
            rejected = false; try { await PasteVisibleAsync(); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.GetCellInput(0, 4, 2) == "" && Document.GetCellInput(0, 9, 0) == "newer", "Stale external preview rejects all destination writes");
            _dialogDriver = w => { w.Close(true); return Task.CompletedTask; }; await PasteVisibleAsync();
            Check(Document.GetCellValue(0, 4, 2) == "00123", "External typed clipboard uses the same visible target mapping");
            _workspace.Editor.SetCell(0, 0, 3, "=A1+1"); _workspace.Editor.SetCell(0, 2, 3, "=A3+2"); Synchronize();
            _grid.SelectRange(new(0, 3, 2, 3), new(0, 3));
            async Task Choose(Window window, int mode)
            {
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                window.GetVisualDescendants().OfType<ComboBox>().Single().SelectedIndex = mode; window.Close(true);
            }
            _dialogDriver = w => Choose(w, 1); await CopyVisibleAsync();
            _grid.SelectRange(new(4, 4, 6, 4), new(4, 4)); _dialogDriver = w => Choose(w, 2); await PasteVisibleAsync();
            Check(Document.GetCellInput(0, 4, 4).Contains("B5") && Document.GetCellInput(0, 6, 4).Contains("B7"), "Formula mode translates each original row to its visible destination");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
