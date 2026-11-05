using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunColorSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        static MenuItem Item(IEnumerable<object?> items, string name) => items.OfType<MenuItem>().Single(m => Equals(m.Header, name));
        async Task Idle()
        {
            for (var i = 0; i < 100; i++)
            {
                await _operation; await Layout();
                if (!_working && !_workspace.IsBusy) { if (!_calculationPending) return; await PerformAsync(CalculatePendingAsync); }
                await Task.Delay(10);
            }
            throw new InvalidOperationException("Color check did not become idle.");
        }
        async Task Click(MenuItem item) { item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await Idle(); }
        async Task Press(Button button) { button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await Idle(); }
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            _workspace.New(); ShowBook(); await Layout();
            var first = Document.Snapshot.ActiveSheetId; await AddSheetAsync(); var second = Document.Snapshot.ActiveSheetId; await Idle();
            Check(_sheets.ItemTemplate!.Build(null) is null, "Tab template accepts null data contexts");
            _grid.Select(new(0, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "原格草稿";
            var target = _sheets.GetVisualDescendants().OfType<ListBoxItem>().Single(c => ((SheetItem)c.DataContext!).Id == first);
            target.RaiseEvent(new ContextRequestedEventArgs()); await Idle();
            Check(_sheets.ContextMenu!.IsOpen && Document.Snapshot.ActiveSheetId == first, "Context request activates the tab under the request");
            Check(Document.Snapshot.GetSheet(second).GetCell(0, 0)?.Value.TextValue == "原格草稿", "Context navigation commits the original sheet's draft");
            var tabs = Item(_sheets.ContextMenu.Items, "标签颜色"); _sheets.ContextMenu.Close();
            await Click(Item(tabs.Items, "红色"));
            Check(Document.Snapshot.GetSheet(first).TabColor == "FFFF0000" && Document.Snapshot.GetSheet(second).TabColor is null, "Tab preset changes only the requested sheet");
            await Click(Item(tabs.Items, "恢复默认颜色")); Check(Document.Snapshot.GetSheet(first).TabColor is null, "Tab reset removes the explicit color");
            await History(false); Check(Document.Snapshot.GetSheet(first).TabColor == "FFFF0000", "Undo reset restores tab color");
            await History(true); Check(Document.Snapshot.GetSheet(first).TabColor is null, "Redo reset removes tab color");
            var revision = Document.Snapshot.Revision;
            _dialogDriver = async window =>
            {
                Check(window is CustomColorDialog, "Tab custom entry uses the reusable dialog"); await Layout();
                Named<TextBox>(window, "ColorHex").Text = "invalid"; await Layout();
                Check(!Named<Button>(window, "ColorConfirm").IsEnabled, "Invalid color disables confirmation");
                Named<TextBox>(window, "ColorHex").Text = "#12ab34"; await Layout();
                Check(Named<Button>(window, "ColorConfirm").IsEnabled, "Valid RGB recovers confirmation");
                var numericText = Named<NumericUpDown>(window, "ColorChannel0").GetVisualDescendants().OfType<TextBox>().Single();
                Check(numericText.Bounds.Width >= 40 && numericText.Text == "18", "Channel numeric editor has visible digits and usable width");
                foreach (var slider in window.GetVisualDescendants().OfType<Slider>())
                {
                    var thumb = slider.GetVisualDescendants().OfType<Thumb>().Single();
                    var position = thumb.TranslatePoint(default, slider)!.Value;
                    Check(position.Y >= 0 && position.Y + thumb.Bounds.Height <= slider.Bounds.Height, "Channel slider thumb fits its visible bounds");
                }
                Check(Document.Snapshot.Revision == revision, "Dialog preview never edits workbook");
                var spectrum = Named<ColorSpectrum>(window, "ColorSpectrum");
                Check(spectrum.SelectedColor == Color.Parse("#FF12AB34"), "Hex input synchronizes the spectrum");
                Named<NumericUpDown>(window, "ColorChannel0").Value = 40; await Layout();
                Check(spectrum.SelectedColor.R == 40, "RGB input synchronizes the spectrum");
                Named<TextBox>(window, "ColorHex").Text = "8000FF00"; await Layout();
                using var pointer = new Pointer(154, PointerType.Mouse, true);
                ulong time = 0;
                Point At(Rect rect, double x, double y) => new(rect.X + rect.Width * x, rect.Y + rect.Height * y);
                Point ScreenPoint(Point point) => spectrum.TranslatePoint(point, window)!.Value;
                void Down(Point point) => spectrum.RaiseEvent(new PointerPressedEventArgs(spectrum, pointer, window, ScreenPoint(point), ++time,
                    new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 1));
                void Move(Point point) => spectrum.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, spectrum, pointer, window, ScreenPoint(point), ++time,
                    new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), KeyModifiers.None));
                void Up(Point point) => spectrum.RaiseEvent(new PointerReleasedEventArgs(spectrum, pointer, window, ScreenPoint(point), ++time,
                    new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased), KeyModifiers.None, MouseButton.Left));
                var square = spectrum.SaturationValueBounds; var hue = spectrum.HueBounds;
                Down(At(hue, 2d / 3, .5)); Up(At(hue, 2d / 3, .5)); await Layout();
                Check(Named<TextBox>(window, "ColorHex").Text == "#800000FF", "Mouse hue selection updates RGB and hex while preserving alpha: " + Named<TextBox>(window, "ColorHex").Text);
                Down(At(square, .5, .5)); await Layout();
                Check(spectrum.SelectedColor.A == 128 && spectrum.SelectedColor.B is >= 127 and <= 128 && spectrum.SelectedColor.R is >= 63 and <= 64,
                    "Mouse spectrum press selects saturation and brightness");
                Move(At(square, 1.5, -.5)); await Layout();
                Check(spectrum.SelectedColor == Color.Parse("#800000FF") && ReferenceEquals(pointer.Captured, spectrum), "Captured drag clamps outside spectrum bounds");
                Up(At(square, 1, 0)); Check(pointer.Captured is null, "Mouse release ends spectrum capture");
                Down(At(square, 1, 1)); Up(At(square, 1, 1));
                Down(At(hue, 1d / 3, .5)); Up(At(hue, 1d / 3, .5));
                Down(At(square, 1, 0)); Up(At(square, 1, 0)); await Layout();
                Check(Named<TextBox>(window, "ColorHex").Text == "#8000FF00", "Hue chosen while black is retained when brightness increases");
                Down(At(square, .25, .25)); pointer.Capture(null); var lost = spectrum.SelectedColor;
                Move(At(square, 1, 1)); Check(spectrum.SelectedColor == lost, "Lost capture stops spectrum updates");
                Named<TextBox>(window, "ColorHex").Text = "#12ab34"; await Layout();
                Check(Document.Snapshot.Revision == revision, "Spectrum drag remains a dialog preview until confirmation");
                await Capture(window, "custom-color");
                Named<Button>(window, "ColorConfirm").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            await Click(Item(tabs.Items, "自定义颜色…"));
            Check(Document.Snapshot.GetSheet(first).TabColor == "FF12AB34", "Custom tab RGB becomes persisted ARGB");
            foreach (var placement in new[] { SheetNavigationPlacement.Bottom, SheetNavigationPlacement.Left, SheetNavigationPlacement.Right })
            {
                await PlaceSheets(placement); await Layout();
                Check(_sheets.GetVisualDescendants().OfType<Border>().Any(b => b.BorderBrush is SolidColorBrush brush && brush.Color == Color.Parse("#FF12AB34") && b.Bounds.Width > 0), "Tab color renders at " + placement);
                foreach (var container in _sheets.GetVisualDescendants().OfType<ListBoxItem>())
                {
                    var text = Named<TextBlock>(container, "SheetTabLabel"); var origin = text.TranslatePoint(default, container)!.Value;
                    Check(Math.Abs(origin.Y - 4) <= 1 && Math.Abs(container.Bounds.Height - origin.Y - text.Bounds.Height - 4) <= 1,
                        "Tab text has approximately 4 DIP top and bottom padding: " + placement + " / " + text.Text);
                }
                if (placement == SheetNavigationPlacement.Bottom)
                    Check(_sheets.Bounds.Height <= 32, "Bottom tab strip fits text without an extra minimum height");
            }
            await PlaceSheets(SheetNavigationPlacement.Bottom); await Idle();
            foreach (var fill in new[] { false, true })
            {
                var button = Named<SplitButton>(this, fill ? "FillColorDropdown" : "FontColorDropdown");
                var primary = Named<Button>(button, "PART_PrimaryButton");
                var secondary = Named<Button>(button, "PART_SecondaryButton");
                var flyout = (MenuFlyout)button.Flyout!;
                var row = fill ? 4 : 0;
                string? CellColor(int r, int c = 0)
                {
                    var snapshot = Document.Snapshot; var style = snapshot.Styles[snapshot.GetSheet(first).GetStyleId(r, c)];
                    return fill ? style.FillColor : style.FontColor;
                }
                bool SwatchIs(string value) => (fill ? _fillColorSwatch : _fontColorSwatch)?.Background is SolidColorBrush brush && brush.Color == Color.Parse("#" + value);
                _grid.Select(new(row, 0)); await Idle();
                await Press(primary);
                Check(!flyout.IsOpen && CellColor(row) == (fill ? "FFFFFF00" : null), "Primary button applies the initial color without opening a menu: " + fill);
                await ApplyStyleAsync(fill ? new(FillColor: "FFFF0000") : new(FontColor: "FFFF0000"));
                var beforeOpen = Document.Snapshot.Revision;
                await Press(secondary);
                Check(Document.Snapshot.Revision == beforeOpen, "Arrow opens the palette without applying a color: " + fill);
                Check(flyout.IsOpen && ColorPalette.Common.All(c => Item(flyout.Items, c.Name).Icon is Border), "Dropdown opens with named color samples: " + fill);
                if (TopLevel.GetTopLevel(Item(flyout.Items, "红色")) is { } popup) await Capture(popup, fill ? "fill-color-menu" : "font-color-menu");
                flyout.Hide(); await Click(Item(flyout.Items, "蓝色"));
                Check((fill ? CurrentStyle.FillColor : CurrentStyle.FontColor) == "FF0000FF", "Preset applies to selected cell: " + fill);
                _grid.SelectRange(new(row + 1, 0, row + 1, 1), new(row + 1, 0)); await Idle();
                Check(CellColor(row + 1) is null && SwatchIs("FF0000FF"), "Moving to unformatted cells retains the last chosen color sample: " + fill);
                await Press(primary);
                Check(!flyout.IsOpen && CellColor(row + 1) == "FF0000FF" && CellColor(row + 1, 1) == "FF0000FF", "Primary button repeats the preset across the new selection: " + fill);
                await History(false);
                Check(CellColor(row + 1) is null && CellColor(row + 1, 1) is null && SwatchIs("FF0000FF"), "Repeated color undoes atomically without changing the remembered color: " + fill);
                await History(true);
                Check(CellColor(row + 1) == "FF0000FF" && CellColor(row + 1, 1) == "FF0000FF", "Repeated color redoes across the selection: " + fill);
                _dialogDriver = async window =>
                {
                    Check(window is CustomColorDialog, "Cell custom entry reuses the same dialog: " + fill); await Layout();
                    Named<TextBox>(window, "ColorHex").Text = "80224466"; await Layout();
                    Named<NumericUpDown>(window, "ColorChannel0").Value = 17; await Layout();
                    Check(Named<TextBox>(window, "ColorHex").Text == "#80114466", "Channel editing synchronizes hex and preserves alpha: " + fill);
                    Named<Button>(window, "ColorConfirm").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                };
                await Click(Item(flyout.Items, "自定义颜色…"));
                Check((fill ? CurrentStyle.FillColor : CurrentStyle.FontColor) == "80114466", "Custom cell color commits: " + fill);
                await History(false); Check((fill ? CurrentStyle.FillColor : CurrentStyle.FontColor) == "FF0000FF", "Cell color undo restores preset: " + fill);
                Check(SwatchIs("80114466"), "Undo retains the confirmed custom color including alpha: " + fill);
                var beforeCancel = Document.Snapshot.Revision;
                _dialogDriver = async window => { await Layout(); Named<TextBox>(window, "ColorHex").Text = "#FFFFFF"; Named<Button>(window, "ColorCancel").RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); };
                await Click(Item(flyout.Items, "自定义颜色…")); Check(Document.Snapshot.Revision == beforeCancel, "Cancel leaves cell color and history unchanged: " + fill);
                await Press(primary);
                Check(!flyout.IsOpen && CellColor(row + 1) == "80114466" && CellColor(row + 1, 1) == "80114466", "Primary repeats the confirmed custom color after undo and cancel: " + fill);
                await Click(Item(flyout.Items, fill ? "无填充" : "自动字体色"));
                Check((fill ? CurrentStyle.FillColor : CurrentStyle.FontColor) is null, "Default cell color can be restored: " + fill);
                await ApplyStyleAsync(fill ? new(FillColor: "FFFF0000") : new(FontColor: "FFFF0000"));
                await Press(primary);
                Check(CellColor(row + 1) is null && CellColor(row + 1, 1) is null, "Primary repeats no fill or automatic font color on formatted cells: " + fill);
                var origin = button.TranslatePoint(default, _ribbon)!.Value;
                Check(button.Bounds.Height <= 28 && origin.Y + button.Bounds.Height <= _ribbon.Bounds.Height, "Color dropdown fits ribbon height: " + fill);
                var icon = button.GetVisualDescendants().OfType<PathIcon>().Single(i => i.Classes.Contains("fluent-ribbon-icon"));
                var iconOrigin = icon.TranslatePoint(default, primary)!.Value;
                Check(icon.Bounds.Size == new Size(20, 20) && iconOrigin.X >= 0 && iconOrigin.Y >= 0 &&
                      iconOrigin.X + 20 <= primary.Bounds.Width && iconOrigin.Y + 20 <= primary.Bounds.Height,
                    "Color icon fits entirely inside the primary button: " + fill);
                Check(!string.IsNullOrWhiteSpace(AutomationProperties.GetName(button)) && ToolTip.GetTip(button) is string { Length: > 0 },
                    "Color split button retains an accessible name and tooltip: " + fill);
                foreach (var part in new[] { primary, secondary })
                {
                    var partOrigin = part.TranslatePoint(default, button)!.Value;
                    Check(part.Bounds.Width > 0 && partOrigin.Y >= 0 && partOrigin.Y + part.Bounds.Height <= button.Bounds.Height,
                        "Both split button targets fit the ribbon: " + fill + " / " + part.Name);
                }
                button.IsEnabled = false; await Layout();
                Check(!primary.IsEffectivelyEnabled && !secondary.IsEffectivelyEnabled && !icon.IsEffectivelyEnabled,
                    "Disabling a color command disables both button targets: " + fill);
                button.IsEnabled = true;
            }
            var fontButton = Named<SplitButton>(this, "FontColorDropdown");
            var fillButton = Named<SplitButton>(this, "FillColorDropdown");
            await Click(Item(((MenuFlyout)fontButton.Flyout!).Items, "红色"));
            await Click(Item(((MenuFlyout)fillButton.Flyout!).Items, "绿色"));
            _grid.Select(new(8, 0)); await Idle();
            await Press(Named<Button>(fontButton, "PART_PrimaryButton"));
            await Press(Named<Button>(fillButton, "PART_PrimaryButton"));
            Check(CurrentStyle.FontColor == "FFFF0000" && CurrentStyle.FillColor == "FF008000", "Font and fill independently remember their last selected colors");
            _grid.BeginEdit(); var presenter = _grid.CellEditor.GetVisualDescendants().OfType<TextPresenter>().Single(); presenter.PreeditText = "ni";
            revision = Document.Snapshot.Revision;
            await SetSheetTabColorAsync(null); await ApplyCellColorAsync(true, "FFFF0000");
            await Press(Named<Button>(fillButton, "PART_PrimaryButton"));
            Check(Document.Snapshot.Revision == revision, "Composition blocks tab and cell color commands");
            presenter.PreeditText = null; _grid.Cancel();
            _grid.Select(new(9, 0)); await Idle(); await Press(Named<Button>(fillButton, "PART_PrimaryButton"));
            Check(CurrentStyle.FillColor == "FF008000", "Blocked color selection does not replace the last successful color");
            await Capture(this, "colors-workspace");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
