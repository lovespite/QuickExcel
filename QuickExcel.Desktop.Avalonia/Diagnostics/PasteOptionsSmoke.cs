using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunPasteOptionsSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            _workspace.New(); ShowBook(); await Layout(); var memory = new MemoryClipboard(); _clipboard = memory;
            var sheet = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand([new(new(sheet, 2, 2), CellContent.FromLiteral(CellValue.FromText("original")))]));
            await ExecuteAsync(new SetCellStyleCommand([new(sheet, 2, 2)], new(Italic: true)));
            var source = new ExternalClipboardRegion(1, 2, [new(CellValue.FromNumber(12), new() { Bold = true }), new(CellValue.FromNumber(34), new())]);
            memory.Value = new(null, SpreadsheetXml: SpreadsheetClipboardCodec.Write(source));
            _grid.Select(new(2, 2)); await PasteAsync("all"); await Layout();
            Check(_pasteIconPopup?.IsOpen == true && _pasteOptionsPopup is null, "Paste shows only the floating icon");
            await Capture(TopLevel.GetTopLevel(_pasteOptionsButton)!, "paste-icon");
            Click(_pasteOptionsButton); await Layout();
            Check(_pasteOptionsPopup?.IsOpen == true && _pasteOptionsPopup.Child is PasteOptionsPanel, "Clicking the icon opens the options panel");
            var panel = (PasteOptionsPanel)_pasteOptionsPopup!.Child!;
            await Capture(TopLevel.GetTopLevel(panel)!, "paste-panel");
            Check(panel.Options.Count == 5 && panel.Options.Values.All(b => b.IsEnabled), "Panel exposes the five supported paste modes");
            memory.Value = new("changed clipboard");
            Click(panel.Options["values"]); await _operation; await Layout();
            var cell = Document.Snapshot.Sheets[0].GetCell(2, 2)!; var style = Document.Snapshot.Styles[cell.Content.StyleId];
            Check(cell.Value == CellValue.FromNumber(12) && style.Italic && !style.Bold, "Values option uses the captured source and restores original formatting");
            Check(_lastPaste is null && _pasteIconPopup is null && _pasteOptionsPopup is null, "Selecting a mode dismisses both panel and icon");
            await History(false);
            Check(Document.Snapshot.Sheets[0].GetCell(2, 2)!.Value == CellValue.FromText("original") && Document.Snapshot.Sheets[0].GetCell(2, 3) is null, "One undo restores the state before the original paste");
            await History(true); Check(Document.Snapshot.Sheets[0].GetCell(2, 2)!.Value == CellValue.FromNumber(12), "Redo restores the chosen mode");

            memory.Value = new(null, SpreadsheetXml: SpreadsheetClipboardCodec.Write(source));
            _grid.Select(new(5, 2)); await PasteAsync("all"); Click(_pasteOptionsButton); await Layout();
            Click(((PasteOptionsPanel)_pasteOptionsPopup!.Child!).Options["transpose"]); await _operation;
            Check(Document.Snapshot.Sheets[0].GetCell(6, 2)!.Value == CellValue.FromNumber(34) && Document.Snapshot.Sheets[0].GetCell(5, 3) is null, "Transpose replaces the old footprint");

            _grid.Select(new(8, 2)); await PasteAsync("all"); Click(_pasteOptionsButton); _grid.BeginEdit();
            Check(_lastPaste is null && _pasteOptionsPopup is null, "Starting a grid draft dismisses the open panel"); _grid.Cancel();
            _grid.Select(new(10, 2)); await PasteAsync("all"); _formula.Focus(); await Layout();
            Check(_lastPaste is null, "Entering the formula bar dismisses the icon");
            _grid.Select(new(12, 2)); await PasteAsync("all"); _grid.Select(new(13, 2)); Check(_lastPaste is null, "Changing selection dismisses stale paste options");
            _grid.Select(new(14, 2)); await PasteAsync("all"); await History(false); Check(_lastPaste is null, "Undo dismisses stale paste options");

            _grid.SelectRange(new(2, 2, 2, 3), new(2, 2)); await CopyAsync(false); _grid.Select(new(16, 2)); await PasteAsync("all");
            Click(_pasteOptionsButton); Click(((PasteOptionsPanel)_pasteOptionsPopup!.Child!).Options["formats"]); await _operation;
            Check((Document.Snapshot.Sheets[0].GetCell(16, 2)?.Value ?? CellValue.Blank) == CellValue.Blank, "Internal format-only replacement restores the destination's original blank value");
            memory.Value = new("=1+2\t7"); _grid.Select(new(18, 2)); await PasteAsync("all"); Click(_pasteOptionsButton);
            panel = (PasteOptionsPanel)_pasteOptionsPopup!.Child!;
            Check(!panel.Options["formats"].IsEnabled, "Plain text disables unavailable source formatting");
            Click(panel.Options["formulas"]); await _operation;
            Check(Document.GetCellInput(0, 18, 2) == "=1+2", "Explicit formula choice retains TSV formula-paste policy");
            _grid.Select(new(20, 2)); await PasteAsync("all"); _workspace.New();
            Check(_lastPaste is null, "Switching documents clears the captured paste");
            await RunExternalClipboardChecksAsync(Check);
            await RunMergedClipboardChecksAsync(Check);
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
