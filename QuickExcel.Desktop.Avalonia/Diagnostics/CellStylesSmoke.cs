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

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunCellStylesSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        static void Click(Control root, string name) => Named<Button>(root, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Idle() { for (var i = 0; i < 100; i++) { await _operation; await Layout(); if (!_working && !_workspace.IsBusy) return; await Task.Delay(10); } throw new TimeoutException(); }
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            _workspace.New(); ShowBook(); Width = 1600; Height = 900; await Layout();
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("金额\n1.25\n2.5\n3.75", new(0, 0, 0, 0), false)); Synchronize(); await FlushEditsAsync();
            _grid.SelectRange(new(1, 0, 3, 0), new(1, 0)); await ApplyStyleAsync(new(NumberFormatId: 4));
            var gallery = Named<Border>(this, "CellStyleGallery");
            Check(gallery.GetVisualDescendants().OfType<Button>().Count(b => b.Name?.StartsWith("CellStyle_") == true) == 10, "Ribbon presents all ten reference tiles");
            Click(this, "CellStyle_good"); await Idle();
            Check(CurrentStyle.FillColor == "FFC6EFCE" && CurrentStyle.FontColor == "FF006100" && CurrentStyle.NumberFormatId == 4, "Good style preserves numeric formatting and applies visible font/fill");
            Check(Document.Snapshot.Sheets[0].GetCell(3, 0)!.Value == CellValue.FromNumber(3.75), "Style selection preserves values");
            await History(false); Check(CurrentStyle.FillColor is null, "Range style application undoes in one step");
            await History(true); Check(CurrentStyle.FillColor == "FFC6EFCE", "Range style application redoes in one step");
            gallery.BringIntoView(); await Layout(); await Capture(gallery, "cell-styles-ribbon");
            var more = Named<Button>(this, "CellStylesMore"); _cellStyleFlyout!.ShowAt(more); await Layout();
            var content = (Control)_cellStyleFlyout.Content!;
            Check(content.GetVisualDescendants().OfType<Button>().Count(b => b.Name?.StartsWith("CellStyle_") == true) == 46, "Expanded gallery contains 46 styles grouped by purpose");
            Check(TopLevel.GetTopLevel(content)!.Bounds.Width >= content.Bounds.Width, "Expanded gallery fits its popup without clipping the fourth column");
            if (TopLevel.GetTopLevel(content) is { } popup) await Capture(popup, "cell-styles-gallery");
            _cellStyleFlyout.Hide();
            Width = 1280; await Layout(); Check(gallery.GetVisualDescendants().OfType<Button>().Count(b => b.IsVisible && b.Name?.StartsWith("CellStyle_") == true) == 6, "Default-width ribbon uses a compact two-row gallery");
            await Capture(this, "cell-styles-default-width"); Width = 1600; await Layout();
            var revision = Document.Snapshot.Revision;
            _dialogDriver = async window => { await Layout(); await Capture(window, "cell-style-new"); Click(window, "CellStyleCancel"); };
            await EditCellStyleAsync(); Check(Document.Snapshot.Revision == revision && _quickStyles.Styles.IsEmpty, "Cancel neither saves a local style nor edits the workbook");
            _dialogDriver = async window =>
            {
                await Layout(); Named<TextBox>(window, "CellStyleName").Text = "常规"; Click(window, "CellStyleConfirm");
                Check(!string.IsNullOrWhiteSpace(Named<TextBlock>(window, "CellStyleError").Text), "Built-in names are rejected inline");
                Named<TextBox>(window, "CellStyleName").Text = "我的绿色填充";
                foreach (var part in new[] { "Number", "Alignment", "Font", "Border", "Protection" }) Named<CheckBox>(window, "CellStylePart" + part).IsChecked = false;
                Click(window, "CellStyleFormat");
                Window? child = null;
                for (var i = 0; i < 50 && child is null; i++) { await Layout(); child = window.OwnedWindows.FirstOrDefault(); if (child is null) await Task.Delay(10); }
                Check(child is NumberFormatDialog, "Style format button opens the full format editor");
                var tabs = Named<TabControl>(child!, "CellFormatTabs"); tabs.SelectedIndex = 3; await Layout();
                var fill = Named<DropDownButton>(child!, "FormatFillColor"); var menu = (MenuFlyout)fill.Flyout!;
                menu.Items.OfType<MenuItem>().First(i => i.Tag is string value && value == "FF008000").RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                await Capture(child!, "cell-style-format-fill"); Click(child!, "FormatConfirm"); await Layout();
                Check(Document.Snapshot.Revision == revision, "Private format previews never mutate the open workbook");
                Click(window, "CellStyleConfirm");
            };
            await EditCellStyleAsync();
            Check(_quickStyles.Styles.Length == 1 && _quickStyles.Styles[0].Parts == CellStyleParts.Fill, "New style saves selected categories to local library");
            var library = new QuickCellStyleStore(Path.Combine(_userDataDirectory, "cell-styles.json")); library.Load();
            Check(library.Error is null && library.Styles.Length == 1 && library.Styles[0].Format.FillColor == "FF008000", "Custom format and category flags survive local reload");
            _workspace.New(); Synchronize(); ShowBook();
            await ApplyStyleAsync(new(FontColor: "FFFF0000", NumberFormatId: 9)); await ApplyStyleAsync(library.Styles[0].Patch);
            Check(CurrentStyle.FillColor == "FF008000" && CurrentStyle.FontColor == "FFFF0000" && CurrentStyle.NumberFormatId == 9, "Local style works in a new workbook and keeps unchecked categories");
            await ApplyStyleAsync(QuickCellStyles.BuiltIn[0].Patch);
            Check(CurrentStyle.FillColor is null && CurrentStyle.NumberFormatId == 0 && CurrentStyle.FontColor is null, "Normal resets all six format categories");
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("名称\n甲\n乙\n丙", new(0, 0, 0, 0), false)); Synchronize();
            await ExecuteAsync(new SetAutoFilterCommand(Document.Snapshot.ActiveSheetId, new(new(0, 0, 3, 0), [new(0, new AutoFilterColumn(values: ["甲", "丙"]))])));
            _grid.SelectRange(new(1, 0, 3, 0), new(1, 0)); Click(this, "CellStyle_bad"); await Idle();
            var sheet = Document.Snapshot.Sheets[0];
            Check(Document.Snapshot.Styles[sheet.GetStyleId(1, 0)].FillColor == "FFFFC7CE" && Document.Snapshot.Styles[sheet.GetStyleId(2, 0)].FillColor is null && Document.Snapshot.Styles[sheet.GetStyleId(3, 0)].FillColor == "FFFFC7CE", "Quick styles skip filtered-out rows");
            _dialogDriver = async window => { await Layout(); Click(window, "DeleteCustomCellStyle"); window.Close((string?)null); };
            await ManageCellStylesAsync(); library.Load(); Check(library.Styles.IsEmpty, "Deleting a local style persists without changing cells");
            Check(Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(1, 0)].FillColor == "FFFFC7CE", "Library management leaves workbook formatting intact");
            await Capture(this, "cell-styles-workspace");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
