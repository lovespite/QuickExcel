using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunMenuSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        async Task Capture(Control control, string name)
        {
            await Layout();
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)control.Bounds.Width, (int)control.Bounds.Height), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        async Task OpenMenu(bool? row = null, int index = 0)
        {
            // Size edits can start the normal deferred calculation while the dialog closes.
            // Route the next input only once that operation has released the host gate.
            for (var attempt = 0; _working && attempt < 1000; attempt++) await Task.Delay(10);
            if (_working) throw new TimeoutException("Menu input remained blocked by a host operation.");
            _grid.ContextMenu?.Close();
            if (row is { } isRow)
            {
                var bounds = _grid.GetCellBounds(isRow ? new(index, 1) : new(1, index))!.Value;
                var point = isRow ? new Point(20 * _grid.Zoom, bounds.Center.Y) : new Point(bounds.Center.X, 18 * _grid.Zoom);
                var canvas = _grid.GetVisualDescendants().OfType<Canvas>().Single();
                canvas.RaiseEvent(new PointerPressedEventArgs(canvas, new Pointer(88, PointerType.Mouse, true), this, canvas.TranslatePoint(point, this)!.Value, 0,
                    new PointerPointProperties(RawInputModifiers.RightMouseButton, PointerUpdateKind.RightButtonPressed), KeyModifiers.None, 1));
            }
            _grid.RaiseEvent(new ContextRequestedEventArgs { RoutedEvent = InputElement.ContextRequestedEvent });
        }
        MenuItem Item(string title) => _grid.ContextMenu!.Items.OfType<MenuItem>().FirstOrDefault(i => i.Header as string == title)
            ?? throw new InvalidOperationException($"Missing {title}; current={_grid.Current}; selection={string.Join(';', _grid.SelectionRanges)}; working={_working}; blocked={_grid.IsContextSelectionBlocked}; composing={_grid.IsComposing}/{_formula.IsComposing}; commit={_grid.IsCommitPending}; menu={string.Join(',', _grid.ContextMenu.Items.OfType<MenuItem>().Select(i => i.Header))}; status={_status.Text}");
        async Task Click(string title)
        {
            var item = Item(title); Check(item.IsEnabled, title + " is enabled");
            item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation; _grid.ContextMenu!.Close(); await Layout();
        }
        void AcceptDialog(string? size = null) => _dialogDriver = async window =>
        {
            await Layout(); if (size is not null) ((TextBox)((FieldsDialog)window).Input("size")).Text = size;
            window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == "确定").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        };
        try
        {
            var memory = new MemoryClipboard(); _clipboard = memory; _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            _workspace.New(true);
            var recentPath = Path.Combine(directory, "年度预算与项目计划.xlsq");
            new TypedWorkbookStorage().Save(recentPath, Document.Snapshot);
            _workspace.New(); _recentFiles.Remember(Path.Combine(directory, "同名文件示例", "年度预算与项目计划.xlsq")); _recentFiles.Remember(recentPath);
            ShowHome(); await Layout();
            var actions = ((StackPanel)_start.Children[0]).Children.OfType<Button>().ToArray();
            Check(actions.Length == 4 && actions.All(b => b.Bounds.Height >= 76 && b.GetVisualDescendants().OfType<Viewbox>().Any()), "Home has four large icon buttons");
            Check(_recent.TranslatePoint(default, _start)!.Value.X > actions[0].Bounds.Right, "Recent files are to the right of the home actions");
            var labels = ((StackPanel)((Button)_recent.Children[0]).Content!).Children.OfType<TextBlock>().ToArray();
            Check(labels[0].Text == Path.GetFileName(recentPath) && labels[0].FontWeight == FontWeight.Bold && labels[0].Foreground == Brushes.Black &&
                labels[1].Text == directory && labels[1].Foreground == Brushes.Gray, "History shows bold black filename above gray directory");
            await Capture(this, "home");
            Width = 760; Height = 480; await Layout(); await Capture(this, "home-minimum");
            Check(_recent.Bounds.Width > 200 && actions.All(b => b.Bounds.Width >= 280), "Minimum window keeps both home columns usable");
            Width = 1280; Height = 820; await Layout();
            ((Button)_recent.Children[0]).RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _operation;
            Check(_book.IsVisible && _workspace.FilePath == recentPath, "History button opens its workbook");
            Check(!_menu.Items.OfType<MenuItem>().Any(i => (i.Header as string)?.Contains("F9") == true), "Main menu no longer contains recalculate F9");
            Check(((MenuItem)_menu.Items[0]!).Items.OfType<Separator>().Count() >= 4, "File menu separates open, save, export, compatibility and close actions");
            _workspace.New(); ShowBook(); var id = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand([new(new(id, 1, 0), CellInput.Parse("12")), new(new(id, 2, 0), CellInput.Parse("34")), new(new(id, 1, 1), CellInput.Parse("=A2*2"))]));
            await ExecuteAsync(new SetAxisMetadataCommand(id, SheetAxis.Row, 1, 2, 36)); await CalculateAsync();
            _grid.SelectRange(new(1, 0, 2, WorkbookLimits.MaxColumns - 1), new(1, 0)); await Layout(); await OpenMenu(true, 2);
            Check(_grid.SelectionRanges.Single().Rows == 2 && _grid.Current.Row == 2, "Right-click within selected row headers retains the entire selection");
            Check(Item("调整行高…") is not null && !Item("插入复制的行").IsEnabled, "Row menu has sizing and disables insert-copy without a row clipboard");
            await Click("复制"); Check(_clipboardAxes?.Count == 2, "Row copy captures both rows without enumerating empty sheet width");
            await OpenMenu(true, 6); await Capture(_grid.ContextMenu!, "row-menu"); await Click("插入复制的行");
            Check(Document.GetCellInput(0, 6, 0) == "12" && Document.GetCellInput(0, 7, 0) == "34" && Document.GetCellValue(0, 6, 1) == "24" && Document.Snapshot.Sheets[0].Rows[6].Height == 36,
                "Row insert-copy preserves values, translated formula and height");
            await History(false); Check(Document.GetCellInput(0, 6, 0) == "", "One undo removes the inserted row block");
            _grid.SelectRange(new(0, 0, WorkbookLimits.MaxRows - 1, 1), new(1, 0)); await OpenMenu(false, 1);
            Check(_grid.SelectionRanges.Single().Columns == 2, "Right-click within selected column headers retains both columns");
            await Click("复制"); await OpenMenu(false, 4); await Capture(_grid.ContextMenu!, "column-menu"); await Click("插入复制的列");
            Check(Document.GetCellInput(0, 1, 4) == "12" && Document.GetCellValue(0, 1, 5) == "24", "Column insert-copy translates formulas and preserves values");
            await History(false); await OpenMenu(false, 1); AcceptDialog("22"); await Click("调整列宽…");
            Check(Document.Snapshot.Sheets[0].Columns[1].Width == 22, "Column width menu commits the entered size");
            await OpenMenu(true, 1); AcceptDialog("44"); await Click("调整行高…");
            Check(Document.Snapshot.Sheets[0].Rows[1].Height == 44, "Row height menu commits the entered size");
            _grid.SelectRange(new(1, 0, 2, WorkbookLimits.MaxColumns - 1), new(1, 0)); await OpenMenu(true, 2); await Click("自动调整行高");
            Check(Document.Snapshot.Sheets[0].Rows[1].Height != 44 && Document.Snapshot.Sheets[0].Rows[2].Height != 36, "Autofit applies to every selected row");
            await History(false); Check(Document.Snapshot.Sheets[0].Rows[1].Height == 44 && Document.Snapshot.Sheets[0].Rows[2].Height == 36, "Multi-row autofit undoes in one step");
            await OpenMenu(true, 1); await Click("隐藏行"); Check(Document.Snapshot.Sheets[0].Rows[1].Hidden, "Row menu hides selected rows");
            await OpenMenu(true, 0); await Click("显示所有行"); Check(!Document.Snapshot.Sheets[0].Rows[1].Hidden, "Row menu restores hidden rows");
            await OpenMenu(false, 1); AcceptDialog(); await Click("删除列"); Check(Document.GetCellInput(0, 1, 1) == "", "Column menu deletes the selected column"); await History(false);
            await OpenMenu(true, 1); AcceptDialog(); await Click("删除行"); Check(Document.GetCellInput(0, 1, 0) == "34", "Row menu deletes the selected row"); await History(false);
            _grid.Select(new(1, 1)); await OpenMenu(); await Capture(_grid.ContextMenu!, "cell-menu");
            Check(Item("设置单元格格式…") is not null && Item("排序与筛选").Items.Count == 4 && Item("选择性粘贴").Items.Count == 4, "Cell menu exposes grouped formatting, data and paste commands");
            await Click("清除内容"); Check(Document.GetCellInput(0, 1, 1) == "", "Cell menu clear command executes"); await History(false);
            _grid.SelectRange(new(1, 0, 1, WorkbookLimits.MaxColumns - 1), new(1, 0)); await CopyAsync(false);
            memory.Value = new("external replacement"); await OpenMenu(true, 6);
            var revision = Document.Snapshot.Revision;
            try { await InsertCopiedAxisAsync(SheetAxis.Row); throw new Exception("Clipboard mismatch unexpectedly accepted"); }
            catch (InvalidOperationException) { Check(Document.Snapshot.Revision == revision, "External clipboard replacement rejects insert-copy without an edit"); }
            _grid.ContextMenu!.Close();
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true;
        (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
