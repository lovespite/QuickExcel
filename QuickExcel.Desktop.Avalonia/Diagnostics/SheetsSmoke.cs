using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
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
    private async Task RunSheetsSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>(); Exception? failure = null;
        void Check(bool ok, string label) { if (!ok) throw new InvalidOperationException(label + " / " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static Button ButtonIn(Control view, string label) => view.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, label));
        static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        static void Choose(MenuItem item) => item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
        static MenuItem MenuEntry(ItemCollection items, string label) => items.OfType<MenuItem>().Single(m => Equals(m.Header, label));
        static void Capture(Control view, string file)
        {
            using var image = new RenderTargetBitmap(new PixelSize((int)view.Bounds.Width, (int)view.Bounds.Height), new Vector(96, 96));
            image.Render(view); image.Save(file, PngBitmapEncoderOptions.Default);
        }
        async Task Idle() { await _operation; await Layout(); }
        ListBoxItem Tab(SheetId id) => _sheets.GetVisualDescendants().OfType<ListBoxItem>().Single(tab => ((SheetItem)tab.DataContext!).Id == id);
        using var mouse = new Pointer(201, PointerType.Mouse, true); ulong time = 0;
        Point At(Control tab, double fraction) => tab.TranslatePoint(new Point(tab.Bounds.Width * fraction, tab.Bounds.Height / 2), this)!.Value;
        void Down(Control tab) => tab.RaiseEvent(new PointerPressedEventArgs(tab, mouse, this, At(tab, .5), ++time,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 1));
        void Move(Point to) => _sheetStrip.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, _sheetStrip, mouse, this, to, ++time,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), KeyModifiers.None));
        void Up(Point to, KeyModifiers keys = KeyModifiers.None) => _sheetStrip.RaiseEvent(new PointerReleasedEventArgs(_sheetStrip, mouse, this, to, ++time,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased), keys, MouseButton.Left));
        async Task Drag(SheetId from, SheetId to, double fraction, KeyModifiers keys = KeyModifiers.None)
        {
            var point = At(Tab(to), fraction); Down(Tab(from)); Move(point);
            Check(_sheetDropMarker.IsVisible && mouse.Captured == _sheetStrip, "Dragging shows the insertion marker and owns capture");
            Up(point, keys); await Idle();
        }
        try
        {
            _workspace.New(); ShowBook(); await Layout();
            var first = Document.Snapshot.ActiveSheetId;
            _grid.BeginEdit(); _grid.CellEditor.Text = "原表内容";
            Click(_addSheetButton); await Idle(); var hidden = Document.Snapshot.ActiveSheetId;
            Check(Document.Snapshot.GetSheet(first).GetCell(0, 0)?.Value.TextValue == "原表内容" && hidden != first, "Trailing plus commits the source draft and creates a sheet");
            Tab(hidden).RaiseEvent(new ContextRequestedEventArgs()); await Idle();
            var context = _sheets.ContextMenu!; var create = MenuEntry(context.Items, "新建工作表");
            context.Close(); Choose(create); await Idle();
            var third = Document.Snapshot.ActiveSheetId;
            Check(Document.Snapshot.Sheets.Length == 3 && third != hidden, "Tab context menu creates a sheet");
            await ExecuteAsync(new SetSheetVisibilityCommand(hidden, SheetVisibility.Hidden)); await Layout();
            Check(Math.Abs(_addSheetButton.TranslatePoint(default, _sheetStrip)!.Value.X - Tab(third).TranslatePoint(new(Tab(third).Bounds.Width, 0), _sheetStrip)!.Value.X) < 4,
                "Plus follows the last tab when the strip fits");
            await Drag(third, first, .1);
            Check(Document.Snapshot.Sheets.Select(s => s.Id).SequenceEqual([third, first, hidden]), "Drag inserts before the first tab using workbook order");
            await Drag(third, first, .9);
            Check(Document.Snapshot.Sheets.Select(s => s.Id).SequenceEqual([first, third, hidden]), "Forward drag inserts after its target");
            await History(false); await Layout();
            Check(Document.Snapshot.Sheets[0].Id == third, "One undo restores sheet order");
            await Drag(first, third, .1, KeyModifiers.Control);
            var copy = Document.Snapshot.Sheets[0];
            Check(copy.Id != first && copy.Name == "Sheet1 (2)" && copy.GetCell(0, 0)?.Value.TextValue == "原表内容" && Document.Snapshot.ActiveSheetId == copy.Id,
                "Ctrl at release copies the source with a distinct name and identity at the insertion marker");
            await History(false); await Layout();
            var revision = Document.Snapshot.Revision;
            var point = At(Tab(first), .5); Down(Tab(third)); Move(point); mouse.Capture(null); Up(point); await Idle();
            Check(Document.Snapshot.Revision == revision && !_sheetDropMarker.IsVisible, "Capture loss cancels a drag");
            Down(Tab(first)); Up(At(Tab(first), .5)); await Idle();
            Check(Document.Snapshot.ActiveSheetId == first && Document.Snapshot.Revision == revision, "Click selects the tab without writing history");

            var secret = Document.AddSheet("内部").Id;
            await ExecuteAsync(new SetSheetVisibilityCommand(secret, SheetVisibility.VeryHidden)); await Layout();
            var mainMenu = MenuEntry(((MenuItem)_menu.Items[2]!).Items, "工作表管理器");
            var contextMenu = MenuEntry(context.Items, "工作表管理器");
            var managerButton = _sheetTools.Children.OfType<Button>().Single(b => AutomationProperties.GetName(b) == "工作表管理器");
            _dialogDriver = async manager =>
            {
                await Layout();
                var list = manager.GetVisualDescendants().OfType<ListBox>().Single();
                var search = manager.GetVisualDescendants().OfType<TextBox>().Single();
                Button Action(string name) => ButtonIn(manager, name);
                SheetId? Selected() => (list.SelectedItem as WorksheetSnapshot)?.Id;
                var active = Document.Snapshot.ActiveSheetId;
                Check(list.ItemCount == 4 && manager.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Text == "深度隐藏"), "Manager displays visible, hidden and very hidden sheets");
                search.Text = "sHEEt"; await Layout();
                Check(Selected() == third, "Changing search selects the first case-insensitive match");
                Click(Action("查找下一个")); await Layout();
                Check(Selected() == first, "Search button advances to the next match");
                search.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter }); await Layout();
                Check(Selected() == hidden && list.ItemCount == 4 && Document.Snapshot.ActiveSheetId == active, "Enter locates a hidden sheet without filtering or activating it");
                Click(Action("查找下一个")); await Layout(); Check(Selected() == third, "Search wraps to the first match");
                search.Text = "absent"; await Layout();
                Check(Selected() is null && list.ItemCount == 4 && !Action("上移").IsEnabled && Action("新建").IsEnabled, "Unmatched search clears selection and disables target actions");
                search.Text = "Sheet2"; await Layout(); Click(Action("上移")); await Layout();
                Check(Selected() == hidden && Document.Snapshot.Sheets[1].Id == hidden, "Move up reorders the hidden target and preserves selection");
                Click(Action("下移")); await Layout(); Check(Document.Snapshot.Sheets[2].Id == hidden, "Move down applies to workbook order");
                var colors = (DropDownButton)Action("设置颜色");
                Choose(MenuEntry(((MenuFlyout)colors.Flyout!).Items, "绿色")); await Layout();
                Check(Document.Snapshot.GetSheet(hidden).TabColor == "FF008000" && Document.Snapshot.GetSheet(active).TabColor is null, "Color applies only to the selected hidden sheet");
                Click(Action("取消隐藏")); await Layout(); Check(Document.Snapshot.GetSheet(hidden).Visibility == SheetVisibility.Visible, "Manager unhides an ordinary hidden sheet");
                Click(Action("隐藏")); await Layout(); Check(Document.Snapshot.GetSheet(hidden).Visibility == SheetVisibility.Hidden, "Manager hides the same sheet again");
                Click(Action("重命名")); await Layout();
                var rename = manager.OwnedWindows.Single(); rename.GetVisualDescendants().OfType<TextBox>().Single().Text = "库存";
                Click(ButtonIn(rename, "确定")); await Layout();
                Check(Document.Snapshot.GetSheet(hidden).Name == "库存" && Selected() == hidden, "Rename retains hidden target identity");
                Click(Action("保护…")); await Layout(); Click(ButtonIn(manager.OwnedWindows.Single(), "确定")); await Layout();
                Check(Document.Snapshot.GetSheet(hidden).Data.Protection?.Sheet == true, "Protection targets the selected sheet");
                Click(Action("复制")); await Layout(); Click(ButtonIn(manager.OwnedWindows.Single(), "确定")); await Layout();
                var hiddenCopy = Document.Snapshot.Sheets.Single(s => s.Name == "库存 (2)");
                Check(Selected() == hiddenCopy.Id && hiddenCopy.Visibility == SheetVisibility.Hidden && Document.Snapshot.ActiveSheetId == active, "Copy selects the new hidden copy while retaining the active tab");
                Click(Action("删除…")); await Layout(); Click(ButtonIn(manager.OwnedWindows.Single(), "确定")); await Layout();
                Check(Document.Snapshot.Sheets.All(s => s.Id != hiddenCopy.Id) && Selected() == secret && !Action("隐藏").IsEnabled, "Delete selects the next entry and very hidden sheets keep their visibility restriction");
                Click(Action("新建")); await Layout(); Check(Selected() == Document.Snapshot.ActiveSheetId && list.ItemCount == 5, "Manager creates and selects a new sheet");
                search.Text = "库存"; await Layout();
                Capture(manager, Path.Combine(directory, "sheet-manager.png")); manager.Close();
            };
            Click(managerButton); await Idle();
            Check(_dialogDriver is null && _status.Text != "对话框专项检查失败。", "Manager opens from the strip");
            foreach (var entry in new[] { mainMenu, contextMenu })
            {
                var opened = false;
                _dialogDriver = dialog => { opened = dialog is SheetManagerDialog; dialog.Close(); return Task.CompletedTask; };
                Choose(entry); await Idle(); Check(opened, "Menu entry opens the manager");
            }
            for (var i = 0; i < 24; i++) Document.AddSheet($"滚动工作表 {i + 1:D2}");
            Synchronize(); await Layout();
            foreach (var placement in new[] { SheetNavigationPlacement.Bottom, SheetNavigationPlacement.Left, SheetNavigationPlacement.Right })
            {
                await PlaceSheets(placement); await Layout();
                var scroll = _sheets.Scroll!; scroll.Offset = default; await Layout();
                var active = Document.Snapshot.ActiveSheetId; revision = Document.Snapshot.Revision;
                Check(!_scrollSheetsBack.IsEnabled && _scrollSheetsForward.IsEnabled, "Overflow arrow states at the start in " + placement);
                Click(_scrollSheetsForward); await Layout();
                Check((VerticalTabs ? scroll.Offset.Y : scroll.Offset.X) > 0 && _scrollSheetsBack.IsEnabled, "Forward arrow scrolls in " + placement);
                Click(_scrollSheetsBack); await Layout();
                Check(Document.Snapshot.ActiveSheetId == active && Document.Snapshot.Revision == revision && !_scrollSheetsBack.IsEnabled, "Scrolling preserves active sheet and history in " + placement);
                scroll.Offset = new Vector(scroll.Extent.Width, scroll.Extent.Height); await Layout();
                Check(!_scrollSheetsForward.IsEnabled && _scrollSheetsBack.IsEnabled, "Forward arrow disables at the end in " + placement);
                if (!VerticalTabs)
                {
                    Click(_addSheetButton); await Idle();
                    var lastTab = Tab(Document.Snapshot.ActiveSheetId);
                    var left = lastTab.TranslatePoint(default, _sheetTabArea)!.Value.X;
                    Check(left >= -1 && left + lastTab.Bounds.Width <= _sheetTabArea.Bounds.Width + 1, "A new active sheet is visible in an overflowing strip");
                    Capture(this, Path.Combine(directory, "sheet-tabs.png"));
                }
            }
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
