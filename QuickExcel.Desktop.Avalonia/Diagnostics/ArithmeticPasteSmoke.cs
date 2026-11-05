using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunArithmeticPasteSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            var memory = new MemoryClipboard(); _clipboard = memory; _workspace.New(); ShowBook();
            _workspace.Editor.SetCells(0, [(0, 0, "8"), (1, 0, "=4+4")]); Synchronize(); await CalculatePendingAsync();
            _grid.SelectRange(new(0, 0, 1, 0), new(0, 0)); memory.Value = new("2");
            Check(((MenuItem)_menu.Items[1]!).Items.OfType<MenuItem>().Any(i => i.Header as string == UiText.Get("运算粘贴…") && i.IsEnabled), "Normal Edit menu exposes arithmetic paste");
            var revision = Document.Snapshot.Revision;
            _dialogDriver = w => { w.Close(false); return Task.CompletedTask; }; await ArithmeticPasteAsync();
            Check(Document.Snapshot.Revision == revision, "Cancel leaves both formula and numeric destinations untouched");
            _dialogDriver = w => { w.Close(true); return Task.CompletedTask; }; await ArithmeticPasteAsync();
            Check(Document.GetCellValue(0, 0, 0) == "10" && Document.GetCellValue(0, 1, 0) == "10" && Document.GetCellInput(0, 1, 0).StartsWith('='), "Scalar operands broadcast while formulas retain their expressions");
            await History(false); Check(Document.GetCellInput(0, 0, 0) == "8" && Document.GetCellInput(0, 1, 0) == "=4+4", "One undo restores the entire arithmetic paste");
            memory.Value = new("2\n0"); _dialogDriver = w => { ((ComboBox)((FieldsDialog)w).Input("operation")).SelectedIndex = 3; w.Close(true); return Task.CompletedTask; };
            var rejected = false; revision = Document.Snapshot.Revision;
            try { await ArithmeticPasteAsync(); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.Snapshot.Revision == revision && Document.GetCellInput(0, 0, 0) == "8", "A later zero divisor rolls back the entire paste");
            memory.Value = new("2"); _dialogDriver = w => { _workspace.Editor.SetCell(0, 8, 0, "newer"); w.Close(true); return Task.CompletedTask; };
            rejected = false; try { await ArithmeticPasteAsync(); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.GetCellInput(0, 0, 0) == "8" && Document.GetCellInput(0, 8, 0) == "newer", "Stale preview preserves newer edits");
            _grid.Select(new(0, 0)); await CopyAsync(false); _grid.Select(new(3, 0));
            _dialogDriver = w => { w.Close(true); return Task.CompletedTask; }; await ArithmeticPasteAsync();
            Check(Document.GetCellValue(0, 3, 0) == "8", "Internal typed copy feeds arithmetic paste");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
