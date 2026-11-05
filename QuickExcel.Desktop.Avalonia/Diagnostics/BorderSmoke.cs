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

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunBorderSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        static void Click(Control root, string name) => Named<Button>(root, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        StyleDefinition Style(int r, int c) => Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(r, c)];
        try
        {
            _workspace.New(); ShowBook(); await Layout(); var sheet = Document.Snapshot.ActiveSheetId;
            _grid.SelectRange(new(1, 1, 3, 3), new(1, 1));
            var button = Named<DropDownButton>(this, "BorderDropdown"); var menu = (MenuFlyout)button.Flyout!;
            menu.ShowAt(button); await Layout(); var presets = menu.Items.OfType<MenuItem>().Where(i => i.Name?.StartsWith("BorderMenu") == true).ToArray();
            Check(presets.Length == 16, "Border menu exposes sixteen range presets");
            if (TopLevel.GetTopLevel(presets[0]) is { } popup) await Capture(popup, "border-menu");
            menu.Hide(); presets.Single(i => i.Name == "BorderMenuTopDoubleBottom").RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation; await Layout();
            Check(Style(1, 1).TopBorder == 1 && Style(3, 3).BottomBorder == 6 && Style(2, 2).BottomBorder == 0, "Menu applies top and double bottom to selection perimeter");
            await History(false); Check(Style(1, 1).TopBorder == 0 && Style(3, 3).BottomBorder == 0, "Menu preset undoes in one step");
            await History(true); Check(Style(3, 3).BottomBorder == 6, "Menu preset redoes");
            var revision = Document.Snapshot.Revision;
            _dialogDriver = async window =>
            {
                await Layout(); Check(Named<TabControl>(window, "CellFormatTabs").SelectedIndex == 2, "Other borders opens the border tab");
                var styles = Named<ListBox>(window, "BorderStyles"); Check(styles.ItemCount == 14, "All thirteen line styles and none are available");
                styles.SelectedIndex = 10;
                var color = Named<DropDownButton>(window, "BorderColor"); var palette = (MenuFlyout)color.Flyout!;
                palette.ShowAt(color); await Layout(); palette.Items.OfType<MenuItem>().Single(i => Equals(i.Header, "蓝色")).RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); palette.Hide();
                Click(window, "BorderPresetOutside"); styles.SelectedIndex = 4; Click(window, "BorderPresetInside");
                styles.SelectedIndex = 1; Click(window, "BorderEdge7");
                Check(Document.Snapshot.Revision == revision, "Border preview and color selection do not mutate workbook");
                await Capture(window, "border-dialog"); Click(window, "FormatCancel");
            };
            await ShowNumberFormatAsync(true, true); Check(Document.Snapshot.Revision == revision, "Cancel discards border draft");
            _dialogDriver = async window =>
            {
                await Layout(); var styles = Named<ListBox>(window, "BorderStyles"); styles.SelectedIndex = 10; Click(window, "BorderPresetOutside");
                styles.SelectedIndex = 4; Click(window, "BorderPresetInside"); styles.SelectedIndex = 1; Click(window, "BorderEdge7");
                Click(window, "FormatConfirm");
            };
            await ShowNumberFormatAsync(true, true);
            Check(Style(1, 1).LeftBorder == 10 && Style(3, 3).BottomBorder == 10 && Style(2, 2).TopBorder == 4 && Style(2, 2).DiagonalDown, "Dialog commits independently styled outside, inside and diagonals");
            await History(false); Check(Style(1, 1).LeftBorder == 0 && Style(3, 3).BottomBorder == 6 && !Style(2, 2).DiagonalDown, "Dialog edits undo as one operation");
            await History(true); revision = Document.Snapshot.Revision;
            _dialogDriver = async window => { await Layout(); Click(window, "FormatConfirm"); };
            await ShowNumberFormatAsync(true, true); Check(Document.Snapshot.Revision == revision, "Reopening and confirming unchanged mixed borders is a no-op");
            _dialogDriver = async window => { await Layout(); Click(window, "BorderPresetNone"); Click(window, "FormatConfirm"); };
            await ShowNumberFormatAsync(true, true); Check(Style(2, 2).TopBorder == 0 && !Style(2, 2).DiagonalDown, "None removes inside and diagonal borders");
            await History(false);
            for (var kind = 1; kind <= 13; kind++)
            {
                var row = 5 + (kind - 1) / 4 * 3; var col = 1 + (kind - 1) % 4 * 2;
                await ExecuteAsync(new SetCellsCommand([new(new(sheet, row, col), CellContent.FromLiteral(CellValue.FromText(BorderEditor.StyleNames[kind])))]));
                await ExecuteAsync(new SetCellBordersCommand([new(sheet, row, col)], CellBorders.Preset(BorderPreset.All, kind, kind % 2 == 0 ? "FF0070C0" : null)));
            }
            await ExecuteAsync(new MergeCellsCommand(sheet, new(1, 5, 3, 7)));
            var merged = Enumerable.Range(1, 3).SelectMany(r => Enumerable.Range(5, 3).Select(c => new CellAddress(sheet, r, c))).ToArray();
            await ExecuteAsync(new SetCellBordersCommand(merged, CellBorders.Preset(BorderPreset.Outside, 6, "FF0070C0") with { Diagonal = new(1, "FF0070C0"), DiagonalUp = true, DiagonalDown = true }));
            _grid.Select(new(0, 0)); await Layout(); await Capture(this, "border-workspace");
            var file = Path.Combine(directory, "borders.xlsx"); new Storage.TypedWorkbookStorage().Export(file, Document.Snapshot);
            var reopened = new Storage.TypedWorkbookStorage().Open(file); Check(reopened.Styles.Any(s => s.DiagonalUp && s.DiagonalDown && s.DiagonalBorder == 1), "Workspace sample saves and reopens with both diagonals");
            await RunMergedBorderSmokeAsync(directory, Check);
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
