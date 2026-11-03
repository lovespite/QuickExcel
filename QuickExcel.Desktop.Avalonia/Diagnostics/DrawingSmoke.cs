using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
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
        }
        catch (Exception ex) { error = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = error is null, checks, error = error?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}




