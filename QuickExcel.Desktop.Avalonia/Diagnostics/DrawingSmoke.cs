using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using Avalonia.VisualTree;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunDrawingSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? error = null;
        void Check(bool ok, string label) { if (!ok) throw new InvalidOperationException(label); checks.Add(label); }
        async Task Ready() { await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background); await _operation; }
        var canvas = (Canvas)((global::Avalonia.Controls.Grid)_grid.Content!).Children[0];
        var pointer = new Pointer(125, PointerType.Mouse, true);
        void Down(Point point) => canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, this, canvas.TranslatePoint(point, this)!.Value, 0,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 1));
        void Move(Point point, KeyModifiers modifiers = KeyModifiers.None) => canvas.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, canvas, pointer, this, canvas.TranslatePoint(point, this)!.Value, 1,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), modifiers));
        void Up(Point point) => canvas.RaiseEvent(new PointerReleasedEventArgs(canvas, pointer, this, canvas.TranslatePoint(point, this)!.Value, 2,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased), KeyModifiers.None, MouseButton.Left));
        void PressKey(Key key) => _grid.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = key });
        async Task Capture(string name)
        {
            await Ready(); using var image = new RenderTargetBitmap(new PixelSize((int)_grid.Bounds.Width, (int)_grid.Bounds.Height), new Vector(96, 96));
            image.Render(_grid); image.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            ShowBook(); var sheet = new WorksheetSnapshot(SheetId.New(), "Objects", view: new() { Zoom = 125, FrozenRows = 1, FrozenColumns = 1, PaneTopLeftCell = new(1, 1) });
            using var pixels = new SKBitmap(40, 20); pixels.Erase(SKColors.Magenta); using var encoded = pixels.Encode(SKEncodedImageFormat.Png, 100);
            var image = new SheetImage(encoded.ToArray(), "png", 0, 0).WithAnchor(DrawingGeometry.FromBounds(sheet, new(180, 100, 160, 90), DrawingAnchorKind.OneCell));
            var chart = SetChartCommand.FromRange(sheet, new(0, 0, 3, 1), WorksheetChartKind.SmoothLine, DrawingGeometry.FromBounds(sheet, new(180, 200, 360, 100), DrawingAnchorKind.TwoCell), "Curve");
            sheet = sheet.WithImages([image]).WithData(new() { Charts = [chart] });
            var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
            var path = Path.Combine(directory, "objects.xlsx"); new TypedWorkbookStorage().Save(path, snapshot); await OpenPathAsync(path); await Ready();
            _grid.ScrollTo(50, 20); await Ready();
            foreach (var target in new[] { new DrawingObjectId(DrawingObjectKind.Image, 0), new DrawingObjectId(DrawingObjectKind.Chart, 0) })
            {
                DrawingAnchor Anchor() => target.Kind == DrawingObjectKind.Image ? Document.Snapshot.Sheets[0].Images[0].Anchor : Document.Snapshot.Sheets[0].Data.Charts[0].Anchor;
                var center = _grid.GetDrawingObjectHandle(target, DrawingHandle.Move)!.Value; var original = Anchor(); var revision = Document.Snapshot.Revision;
                Down(center); Up(center); await Ready(); Check(_grid.SelectedDrawingObject == target && Document.Snapshot.Revision == revision, $"{target.Kind} click selects without dirtying workbook (selected={_grid.SelectedDrawingObject}, expected={target}, revision={Document.Snapshot.Revision}/{revision}, position={center})");
                Down(center); Move(center + new Vector(25, 15)); Check(Document.Snapshot.Revision == revision, target.Kind + " pointer preview does not write cells or history");
                Up(center + new Vector(25, 15)); await Ready(); var moved = Anchor(); var b = original.Bounds(Document.Snapshot.Sheets[0]); var m = moved.Bounds(Document.Snapshot.Sheets[0]);
                Check(Math.Abs(m.X - b.X - 20) < .01 && Math.Abs(m.Y - b.Y - 12) < .01, target.Kind + " moves at 125% zoom in scrolled frozen panes");
                await History(false); Check(Anchor() == original, target.Kind + " move undo restores exact anchor"); await History(true);
                var corner = _grid.GetDrawingObjectHandle(target, DrawingHandle.BottomRight)!.Value;
                Down(corner); Move(corner + new Vector(35, 20)); Up(corner + new Vector(35, 20)); await Ready(); Check(Anchor().Bounds(Document.Snapshot.Sheets[0]).Width > m.Width, target.Kind + " resize handle commits dimensions");
                var knob = _grid.GetDrawingObjectHandle(target, DrawingHandle.Rotate)!.Value; center = _grid.GetDrawingObjectHandle(target, DrawingHandle.Move)!.Value;
                var rotated = DrawingObjectInteraction.Rotate(new(knob.X, knob.Y), new(center.X, center.Y), 30); var end = new Point(rotated.X, rotated.Y);
                Down(knob); Move(end, KeyModifiers.Shift); Up(end); await Ready(); Check(Math.Abs(Anchor().Rotation - 30) < .001, target.Kind + " rotation handle supports 15-degree snapping");
                await Capture(target.Kind + "-selected");
                original = Anchor(); center = _grid.GetDrawingObjectHandle(target, DrawingHandle.Move)!.Value; revision = Document.Snapshot.Revision;
                Down(center); Move(center + new Vector(20, 20)); PressKey(Key.Escape); Up(center); await Ready(); Check(Anchor() == original && Document.Snapshot.Revision == revision, target.Kind + " Escape cancels preview without history");
                Down(center); Move(center + new Vector(20, 20)); pointer.Capture(null); await Ready(); Check(Anchor() == original, target.Kind + " capture loss cancels gesture");
                Down(center); Move(center + new Vector(20, 20));
                await ExecuteAsync(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 10, 10), CellContent.FromLiteral(CellValue.FromNumber((int)target.Kind + 1)))]));
                await Ready(); revision = Document.Snapshot.Revision; Up(center + new Vector(20, 20)); await Ready();
                Check(Anchor() == original && Document.Snapshot.Revision == revision && !_grid.IsDrawingObjectDragging, target.Kind + " concurrent document change cancels stale drag");
            }
            var beforeDelete = Document.Snapshot; PressKey(Key.Delete); await Ready(); Check(Document.Snapshot.Sheets[0].Data.Charts.IsEmpty && Document.Snapshot.Sheets[0].Images.Length == 1, "Delete removes selected object only");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.Charts[0].Anchor == beforeDelete.Sheets[0].Data.Charts[0].Anchor, "Object deletion is undoable");
            var output = Path.Combine(directory, "objects-edited.xlsx"); new TypedWorkbookStorage().Save(output, Document.Snapshot); var read = new TypedWorkbookStorage().Open(output);
            Check(read.Sheets[0].Images[0].Anchor.Rotation == 30 && read.Sheets[0].Data.Charts[0].Anchor.Rotation == 30, "Image/chart transforms survive XLSX save and reopen");
            await ExecuteAsync(new SetProtectionCommand(Document.Snapshot.ActiveSheetId, new() { Objects = true }));
            var imageTarget = new DrawingObjectId(DrawingObjectKind.Image, 0); var protectedAnchor = Document.Snapshot.Sheets[0].Images[0].Anchor;
            var hit = _grid.GetDrawingObjectHandle(imageTarget, DrawingHandle.Move)!.Value; Down(hit); Move(hit + new Vector(40, 30)); Up(hit); await Ready();
            Check(Document.Snapshot.Sheets[0].Images[0].Anchor == protectedAnchor && !_grid.CanEditDrawingObjects, "Protected objects remain selectable and cannot be transformed");
            await Capture("protected-objects");
            // Chart selection, filter popup, history and file persistence through the actual desktop handlers.
            var cells = new List<KeyValuePair<CellPosition, CellSnapshot>>();
            void Cell(int r, int c, CellValue value) => cells.Add(new(new(r, c), new(CellContent.FromLiteral(value))));
            Cell(0, 0, CellValue.FromText("货号")); Cell(0, 1, CellValue.FromText("价格")); Cell(0, 2, CellValue.FromText("数量"));
            for (var r = 1; r <= 9; r++) { Cell(r, 0, CellValue.FromText("X" + (126 + r))); Cell(r, 1, CellValue.FromNumber(r == 9 ? 3556.11 : r * 80)); Cell(r, 2, CellValue.FromNumber(r * 5)); }
            sheet = new WorksheetSnapshot(SheetId.New(), "图表筛选", cells);
            chart = SetChartCommand.FromRange(sheet, new(0, 0, 9, 2), WorksheetChartKind.Column,
                DrawingGeometry.FromBounds(sheet, new(355, 10, 580, 340), DrawingAnchorKind.OneCell), "价格与数量");
            snapshot = new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet.WithData(sheet.Data with { Charts = [chart] })]);
            path = Path.Combine(directory, "chart-filters.xlsx"); new TypedWorkbookStorage().Save(path, snapshot); await OpenPathAsync(path); await Ready();
            var chartTarget = new DrawingObjectId(DrawingObjectKind.Chart, 0); var chartCenter = _grid.GetDrawingObjectHandle(chartTarget, DrawingHandle.Move)!.Value;
            Down(chartCenter); Up(chartCenter); await Ready(); var chartRevision = Document.Snapshot.Revision;
            Check(_grid.SelectedChartSources.Count == 5 && _grid.ChartFilterButtonBounds is not null && !Document.IsDirty, "Chart selection marks complete names/categories/values and exposes filter button without dirtying");
            await Capture("chart-sources");
            _chartFilterDriver = async panel =>
            {
                panel.SeriesItems[1].Selected = false; panel.CategoryItems[1].Selected = false;
                Check(Document.Snapshot.Revision == chartRevision, "Chart filter edits remain a draft until Apply");
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                using var menuImage = new RenderTargetBitmap(new PixelSize((int)panel.Bounds.Width, (int)panel.Bounds.Height), new Vector(96, 96)); menuImage.Render(panel); menuImage.Save(Path.Combine(directory, "chart-filter-panel.png"), PngBitmapEncoderOptions.Default);
                panel.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            canvas.Children.OfType<Button>().Single().RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await Ready();
            var filteredChart = Document.Snapshot.Sheets[0].Data.Charts[0];
            Check(filteredChart.Series[1].IsFiltered && filteredChart.Series.All(s => s.HiddenPoints.SequenceEqual([1])) && Document.Snapshot.Revision == chartRevision + 1,
                "Chart filter button applies series and category choices as one document command");
            Check(ChartData.Read(Document.Snapshot, filteredChart).Single().Values.Length == 8 && Document.Snapshot.Sheets[0].Cells.Count == cells.Count, "Chart filter removes plotted data without changing worksheet cells");
            await Capture("chart-filtered");
            chartRevision = Document.Snapshot.Revision;
            _chartFilterDriver = panel =>
            {
                Check(panel.SeriesItems[1].Selected == false && panel.CategoryItems[1].Selected == false, "Reopening chart filter restores saved checkbox choices");
                panel.CategoryItems[0].Selected = false; panel.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape }); return Task.CompletedTask;
            };
            await ChartFilterAsync(); Check(Document.Snapshot.Revision == chartRevision, "Escape discards unapplied chart filter changes");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.Charts[0].Series.All(s => !s.IsFiltered && s.HiddenPoints.IsEmpty), "Undo restores all chart data");
            await History(true); Check(Document.Snapshot.Sheets[0].Data.Charts[0].Series[1].IsFiltered, "Redo restores chart filtering");
            foreach (var ext in new[] { ".xlsx", ".xlsq" })
            {
                var saved = Path.Combine(directory, "chart-filter-saved" + ext); new TypedWorkbookStorage().Save(saved, Document.Snapshot);
                await OpenPathAsync(saved); await Ready(); chartCenter = _grid.GetDrawingObjectHandle(chartTarget, DrawingHandle.Move)!.Value; Down(chartCenter); Up(chartCenter); await Ready();
                _chartFilterDriver = panel => { Check(panel.SeriesItems[1].Selected == false && panel.CategoryItems[1].Selected == false, ext + " reopen restores chart filter panel"); panel.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape }); return Task.CompletedTask; };
                await ChartFilterAsync();
            }
            _chartFilterDriver = async panel =>
            {
                var search = panel.GetVisualDescendants().OfType<TextBox>().Single(t => t.Name == "ChartFilterSearch"); search.Text = "X127";
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                panel.GetVisualDescendants().OfType<CheckBox>().Single(c => c.Name == "ChartCategoriesAll").IsChecked = false;
                Check(panel.CategoryItems[0].Selected == false && panel.CategoryItems[2].Selected == true, "Search select-all only changes matching chart categories");
                panel.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape });
            };
            await ChartFilterAsync();
            _chartFilterDriver = async panel =>
            {
                panel.CategoryItems[0].Selected = false;
                await ExecuteAsync(new SetCellsCommand([new(new(Document.Snapshot.ActiveSheetId, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(77)))]));
                Check(_chartFilterPopup?.IsOpen == false, "Source edit closes stale chart filter panel");
            };
            await ChartFilterAsync(); Check(Document.Snapshot.Sheets[0].Data.Charts[0].Series[0].HiddenPoints.SequenceEqual([1]), "Stale chart filter draft cannot overwrite current state");
            PressKey(Key.Escape); await Ready(); Check(_grid.SelectedChartSources.Count == 0 && _grid.ChartFilterButtonBounds is null, "Deselecting chart removes source marks and filter button");
            chartCenter = _grid.GetDrawingObjectHandle(chartTarget, DrawingHandle.Move)!.Value; Down(chartCenter); Up(chartCenter); await Ready();
            await ExecuteAsync(new SetProtectionCommand(Document.Snapshot.ActiveSheetId, new() { Objects = true }));
            _chartFilterDriver = panel => { Check(!panel.ApplyButton.IsEnabled, "Protected chart filter is inspectable but cannot apply"); panel.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape }); return Task.CompletedTask; };
            await ChartFilterAsync();
        }
        catch (Exception ex) { error = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = error is null, checks, error = error?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}




