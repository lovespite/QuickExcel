using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunTextRotationSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        async Task Idle() { await _operation; await Layout(); }
        async Task<string> Capture(Control control, string name)
        {
            await Layout(); var path = Path.Combine(directory, name + ".png");
            using var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            image.Render(control); image.Save(path, PngBitmapEncoderOptions.Default); return path;
        }
        try
        {
            int[] rotations = [45, 135, 255, 90, 180, 0, 27, 117];
            var style = new StyleDefinition { FontSize = 18, Bold = true, Horizontal = 2, Vertical = 1 };
            var cells = new Dictionary<CellPosition, CellSnapshot>();
            for (var column = 0; column < rotations.Length; column++)
            {
                cells[new(0, column)] = new(CellContent.FromLiteral(CellValue.FromText(column < 6 ? UiText.Get(TextRotationOptions.Presets[column].Name) : column == 6 ? "+27°" : "−27°")));
                cells[new(1, column)] = new(CellContent.FromLiteral(CellValue.FromText("旋转 Abc"), column + 1));
            }
            cells[new(3, 0)] = new(CellContent.FromLiteral(CellValue.FromText("合并单元格 Abc"), 1));
            var sheet = new WorksheetSnapshot(SheetId.New(), "Rotation", cells,
                rows: [new(0, new(Height: 28)), new(1, new(Height: 200)), new(3, new(Height: 160))],
                columns: Enumerable.Range(0, 8).Select(column => new KeyValuePair<int, ColumnMetadata>(column, new(Width: 17))), mergedCells: ["A4:C4"]);
            var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet],
                styles: new[] { new StyleDefinition() }.Concat(rotations.Select(rotation => style with { TextRotation = rotation })));
            var storage = new TypedWorkbookStorage(); var source = Path.Combine(directory, "rotation.xlsq"); storage.Save(source, snapshot);
            await OpenPathAsync(source); await Layout();
            var dropdown = Named<DropDownButton>(this, "TextRotationDropdown"); var menu = (MenuFlyout)dropdown.Flyout!;
            _grid.Select(new(1, 0)); menu.ShowAt(dropdown); await Layout();
            var presets = menu.Items.OfType<MenuItem>().Where(item => item.Name?.StartsWith("TextRotation", StringComparison.Ordinal) == true).ToArray();
            Check(presets.Length == 6 && presets.Single(item => item.IsChecked).Name == "TextRotation45", "Orientation menu has five presets, horizontal reset and current selection");
            if (TopLevel.GetTopLevel(presets[0]) is { } popup) await Capture(popup, "rotation-menu"); menu.Hide();
            await Capture(this, "rotation-workspace");
            Check(_grid.MeasureAutoFit(SheetAxis.Column, 2) < _grid.MeasureAutoFit(SheetAxis.Column, 5), "Auto-fit measures stacked glyph width instead of horizontal text width");

            SKColor[]? editorPixels = null;
            foreach (var rotation in new[] { 0, 45, 135, 255, 90, 180 })
            {
                _grid.SelectRange(new(1, 0, 1, 1), new(1, 0));
                presets.Single(item => item.Name == "TextRotation" + rotation).RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await Idle();
                Check(Enumerable.Range(0, 2).All(column => Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(1, column)].TextRotation == rotation), "Menu applies rotation to whole selection: " + rotation);
                _grid.Select(new(1, 0)); Check(_grid.BeginEdit(), "Rotated cell enters native editor: " + rotation);
                var editor = _grid.CellEditor; editor.SelectionStart = editor.SelectionEnd = editor.CaretIndex = 0; _grid.Focus(); await Layout();
                using var pixels = SKBitmap.Decode(await Capture(editor, "rotation-editor-" + rotation));
                if (editorPixels is null) editorPixels = pixels.Pixels;
                else Check(editorPixels.SequenceEqual(pixels.Pixels), "Editor stays horizontal with unchanged text pixels: " + rotation);
                Check(editor.Text == "旋转 Abc", "Editing preserves text: " + rotation); _grid.Cancel();
            }
            await History(false); Check(CurrentStyle.TextRotation == 90, "Range rotation undoes in one step");
            await History(true); Check(CurrentStyle.TextRotation == 180, "Range rotation redoes in one step");
            _dialogDriver = async window =>
            {
                await Layout(); Check(Named<TabControl>(window, "CellFormatTabs").SelectedIndex == 1, "Alignment menu opens alignment tab directly");
                var degrees = Named<NumericUpDown>(window, "RotationDegrees"); Check(degrees.Value == -90, "Dialog restores signed angle");
                degrees.Value = 12.5m;
                Named<Button>(window, "FormatConfirm").RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await Layout();
                Check(!string.IsNullOrEmpty(Named<TextBlock>(window, "FormatError").Text), "Fractional angle stays in dialog with validation error");
                degrees.Value = -27; await Layout(); await Capture(window, "rotation-custom");
                Named<Button>(window, "FormatConfirm").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            menu.Items.OfType<MenuItem>().Last().RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await Idle();
            Check(CurrentStyle.TextRotation == 117, "Custom clockwise angle commits as portable style");
            var revision = Document.Snapshot.Revision;
            _dialogDriver = async window =>
            {
                await Layout(); Named<CheckBox>(window, "RotationStacked").IsChecked = true; await Layout();
                Check(!Named<NumericUpDown>(window, "RotationDegrees").IsEnabled, "Stacked mode disables angle input");
                Named<Button>(window, "FormatCancel").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            await ShowNumberFormatAsync(true, alignmentTab: true); Check(Document.Snapshot.Revision == revision, "Cancelling orientation leaves workbook unchanged");
            _grid.Select(new(3, 1)); Check(_grid.BeginEdit(), "Merged rotated cell enters editor");
            Check(_grid.Current == new CellPosition(3, 0) && _grid.CellEditor.Text == "合并单元格 Abc" && _grid.CellEditor.RenderTransform is null,
                "Merged editor retains anchor and normal text direction"); _grid.Cancel();
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var path = Path.Combine(directory, "saved." + extension);
                if (extension == "xlsx") storage.Export(path, Document.Snapshot); else storage.Save(path, Document.Snapshot);
                var reopened = storage.Open(path);
                Check(reopened.Styles[reopened.Sheets[0].GetStyleId(1, 0)].TextRotation == 117 && reopened.Sheets[0].GetCell(1, 0)!.Value.TextValue == "旋转 Abc", "Save/reopen preserves custom angle and text: " + extension);
            }
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
