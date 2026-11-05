using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunPageBreakSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        async Task<string> Capture(string name)
        {
            await Layout(); var file = Path.Combine(directory, name + ".png");
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)_grid.Bounds.Width, (int)_grid.Bounds.Height), new Vector(96, 96));
            bitmap.Render(_grid); bitmap.Save(file, PngBitmapEncoderOptions.Default); return file;
        }
        bool Marker(string file, bool row, int index)
        {
            var bounds = _grid.GetCellBounds(row ? new(index, 1) : new(1, index))!.Value;
            var zoom = _grid.Zoom;
            var region = row ? new Rect(3 * zoom, bounds.Y + 2 * zoom, 14 * zoom, 8 * zoom)
                : new Rect(bounds.X + 2 * zoom, 3 * zoom, 8 * zoom, 14 * zoom);
            using var bitmap = SKBitmap.Decode(file);
            for (var y = (int)region.Top; y < region.Bottom; y++)
            for (var x = (int)region.Left; x < region.Right; x++)
            { var color = bitmap.GetPixel(x, y); if (color.Blue > 190 && color.Red < 80 && color.Green is > 60 and < 150) return true; }
            return false;
        }
        async Task BreakDialog(int axis, int action)
        {
            _dialogDriver = async window =>
            {
                await Layout(); var dialog = (FieldsDialog)window;
                ((ComboBox)dialog.Input("axis")).SelectedIndex = axis;
                ((ComboBox)dialog.Input("action")).SelectedIndex = action;
                window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == "确定").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            try { await PageBreaksAsync(); } finally { _dialogDriver = null; }
        }
        void RightClickHeader(bool row, int index)
        {
            var bounds = _grid.GetCellBounds(row ? new(index, 1) : new(1, index))!.Value;
            var point = row ? new Point(20 * _grid.Zoom, bounds.Center.Y) : new Point(bounds.Center.X, 18 * _grid.Zoom);
            var canvas = _grid.GetVisualDescendants().OfType<Canvas>().Single();
            canvas.RaiseEvent(new PointerPressedEventArgs(canvas, new Pointer(82, PointerType.Mouse, true), this, canvas.TranslatePoint(point, this)!.Value, 0,
                new PointerPointProperties(RawInputModifiers.RightMouseButton, PointerUpdateKind.RightButtonPressed), KeyModifiers.None, 1));
            _grid.RaiseEvent(new ContextRequestedEventArgs { RoutedEvent = InputElement.ContextRequestedEvent });
        }
        MenuItem Item(string title) => _grid.ContextMenu!.Items.OfType<MenuItem>().Single(i => i.Header as string == title);
        async Task ClickMenu(string title, bool expectStale = false)
        {
            var item = Item(title); Check(item.IsEnabled, title + " is enabled");
            if (expectStale) _status.Text = "";
            item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            try { await _operation; }
            finally { _grid.ContextMenu!.Close(); }
            if (expectStale) Check(_status.Text == UiText.Get("文档已发生变化，请重新提交操作。"), "Stale context-menu command reports revision conflict");
            await Layout();
        }
        try
        {
            ShowBook();
            var seed = Path.Combine(directory, "source.xlsq");
            new TypedWorkbookStorage().Save(seed, new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "分页检查")]));
            await OpenPathAsync(seed); await Layout();
            var id = Document.Snapshot.ActiveSheetId;
            var original = new WorksheetPrinting { RowBreaks = [new(8, 0, PrintRange.MaxColumns - 1)], HeaderFooter = new() { OddHeader = "分页检查" } };
            await ExecuteAsync(new SetPrintingCommand(id, original));
            _grid.Select(new(5, 2)); await BreakDialog(0, 0); await BreakDialog(1, 0);
            Check(Document.Snapshot.Sheets[0].Printing.RowBreaks.Any(b => b.Id == 5 && b.Manual), "Dialog inserts a manual break before row 6");
            var file = await Capture("page-breaks");
            Check(Marker(file, true, 5) && Marker(file, false, 2), "Manual row and column breaks have blue header markers");
            Check(!Marker(file, true, 8) && !Marker(file, true, 4), "Automatic breaks and adjacent rows have no manual marker");
            var rowBounds = _grid.GetCellBounds(new(5, 1))!.Value;
            Check(_grid.HeaderAnnotation(new(20, rowBounds.Center.Y))?.Contains("第 6 行前") == true, "Header tooltip names the exact boundary and delete action");
            await FreezeAsync(2, 1); await ZoomAsync(150); _grid.ScrollTo(30, 15);
            file = await Capture("page-breaks-zoomed-frozen");
            Check(Marker(file, true, 5) && Marker(file, false, 2), "Markers track scrolled, frozen, zoomed headers");
            _grid.Select(new(0, 0)); RightClickHeader(true, 5);
            Check(_grid.Current.Row == 5, $"Right-click targets the clicked row header instead of the previous selection (row index {_grid.Current.Row})");
            await ClickMenu("删除行分页符");
            Check(!Document.Snapshot.Sheets[0].Printing.RowBreaks.Any(b => b.Id == 5) && Document.Snapshot.Sheets[0].Printing.ColumnBreaks.Single().Id == 2,
                "Deleting a row break preserves the column break");
            file = await Capture("page-break-deleted"); Check(!Marker(file, true, 5), "Deleting a break invalidates the cached header marker");
            await History(false); file = await Capture("page-break-undo"); Check(Marker(file, true, 5), "Undo restores the marker immediately");
            await History(true); file = await Capture("page-break-redo"); Check(!Marker(file, true, 5), "Redo removes the marker immediately");
            RightClickHeader(false, 2); await ClickMenu("删除列分页符");
            Check(Document.Snapshot.Sheets[0].Printing.ContentEquals(original), "Deleting both manual breaks preserves automatic breaks and other print settings");
            foreach (var extension in new[] { ".xlsx", ".xlsq" })
            {
                var path = Path.Combine(directory, "deleted" + extension); new TypedWorkbookStorage().Save(path, Document.Snapshot);
                Check(new TypedWorkbookStorage().Open(path).Sheets[0].Printing.ContentEquals(original), "Deleted page breaks stay deleted after reopening " + extension);
            }
            RightClickHeader(true, 5); Check(!Item("删除行分页符").IsEnabled, "Delete is disabled at a row without a break");
            await ClickMenu("在当前行前插入分页符");
            RightClickHeader(true, 5); _workspace.Editor.SetCell(0, 12, 0, "concurrent"); Synchronize();
            await ClickMenu("删除行分页符", expectStale: true);
            Check(Document.Snapshot.Sheets[0].Printing.RowBreaks.Any(b => b.Id == 5), "Stale context-menu deletion rejects without losing the break");
            _grid.Select(new(0, 0)); _grid.RaiseEvent(new ContextRequestedEventArgs { RoutedEvent = InputElement.ContextRequestedEvent });
            Check(!Item("在当前行前插入分页符").IsEnabled && !Item("在当前列前插入分页符").IsEnabled, "First row and column reject insertion"); _grid.ContextMenu!.Close();
            await BreakDialog(0, 2);
            Check(Document.Snapshot.Sheets[0].Printing.ContentEquals(original), "Clear-all removes manual breaks while retaining automatic breaks");
            await ZoomAsync(100); _grid.ScrollTo(0, 0);
            await ExecuteAsync(new SetAxisMetadataCommand(id, SheetAxis.Row, 3, 1, hidden: true));
            _grid.Select(new(5, 2)); await BreakDialog(0, 0);
            file = await Capture("page-break-hidden-rows"); Check(Marker(file, true, 5), "Hidden preceding rows do not shift the break marker to another row");
            var beforeSelection = _grid.BeforeSelection;
            try
            {
                _grid.BeforeSelection = () => false; RightClickHeader(true, 6);
                Check(!_grid.ContextMenu!.IsOpen && Document.Snapshot.Sheets[0].Printing.RowBreaks.Any(b => b.Id == 5),
                    "A rejected pending edit cannot expose deletion for the previously selected row");
            }
            finally { _grid.BeforeSelection = beforeSelection; }
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
