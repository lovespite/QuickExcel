using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckReadOnlyWorkbookAsync(string directory, Action<bool, string> check)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        var source = Path.Combine(directory, "read-only-source.xlsq");
        var first = new WorksheetSnapshot(SheetId.New(), "First", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("只读原件"))))]);
        var second = new WorksheetSnapshot(SheetId.New(), "Second");
        new TypedWorkbookStorage().Save(source, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [first, second]));
        var original = File.ReadAllBytes(source); var originalFiles = _files; var originalClipboard = _clipboard;
        if (OperatingSystem.IsWindows()) File.SetAttributes(source, FileAttributes.ReadOnly);
        try
        {
            var dialogs = new SmokeFiles(); _files = dialogs; _clipboard = new MemoryClipboard();
            await OpenPathAsync(source, readOnly: !OperatingSystem.IsWindows()); await Layout();
            check(Document.IsReadOnly && !Document.IsDirty && _readOnlyNotice.IsVisible && Title!.Contains(UiText.Get(" [只读]")),
                "A source without write permission opens read-only with a visible banner and title");
            check(_formula.IsReadOnly && _grid.IsReadOnly && !_grid.BeginEdit() && !_grid.CanEditDrawingObjects && !_addSheetButton.IsEnabled,
                "Read-only grid and formula bar reject editing, drawing manipulation and adding sheets");
            check(_mutationControls.All(control => !control.IsEnabled) && _readOnlyControls.All(control => !control.IsEnabled),
                "Read-only mode disables formatting, structure, calculation and script controls");
            var fileMenu = (MenuItem)_menu.Items[0]!; var editMenu = (MenuItem)_menu.Items[1]!;
            MenuItem Item(MenuItem menu, string title) => menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get(title)));
            check(new[] { "另存为", "导出值快照", "打印预览 / PDF", "兼容性报告" }.All(title => Item(fileMenu, title).IsEnabled) &&
                Item(editMenu, "复制  Ctrl+C").IsEnabled && Item(editMenu, "查找与替换  Ctrl+F").IsEnabled,
                "Read-only workbooks retain Save As, export, print, compatibility, copy and find");
            check(!Item(editMenu, "粘贴  Ctrl+V").IsEnabled && !Item(editMenu, "剪切  Ctrl+X").IsEnabled && !((MenuItem)_menu.Items[2]!).IsEnabled,
                "Read-only menus visibly disable paste, cut and worksheet mutations");
            await CopyAsync(false); check(((MemoryClipboard)_clipboard).Value.Text == "只读原件", "Copy works from read-only cells");
            var revision = Document.Snapshot.Revision;
            Item(editMenu, "粘贴  Ctrl+V").RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation;
            foreach (var key in new[] { Key.V, Key.X, Key.D, Key.Z })
                _grid.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = key, KeyModifiers = KeyModifiers.Control });
            _grid.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Delete });
            await _operation;
            check(Document.Snapshot.Revision == revision && !Document.IsDirty && Document.GetCellInput(0, 0, 0) == "只读原件",
                "Routed disabled menu clicks and edit shortcuts cannot modify a read-only workbook");
            BuildGridContextMenu(_grid.ContextMenu!, null);
            check(_grid.ContextMenu!.Items.OfType<MenuItem>().All(item => ((string)item.Header!).StartsWith(UiText.Get("复制"))),
                "Read-only cell context menus offer copy commands without editing commands");
            await NavigateToAsync(second.Id, new(2, 3, 2, 3));
            check(Document.Snapshot.ActiveSheetId == second.Id && _grid.Current == new CellPosition(2, 3) && !Document.IsDirty,
                "Read-only navigation can change sheets and selection without dirtying the document");
            await NavigateToAsync(first.Id, new(0, 0, 0, 0));
            _dialogDriver = async window =>
            {
                await Layout();
                var buttons = window.GetVisualDescendants().OfType<Button>().ToArray();
                check(buttons.Where(button => Equals(button.Content, UiText.Get("替换所选")) || Equals(button.Content, UiText.Get("全部替换"))).All(button => !button.IsEnabled),
                    "Find keeps replacement buttons disabled in read-only mode");
                window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == UiText.Get("查找")).Text = "只读";
                buttons.Single(button => Equals(button.Content, UiText.Get("查找全部"))).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                var results = window.GetVisualDescendants().OfType<ListBox>().Single();
                for (var attempt = 0; results.Items.Count == 0 && attempt < 100; attempt++) await Task.Delay(10);
                check(results.Items.Count == 1, "Find returns matches in a read-only workbook"); window.Close();
            };
            await SearchAsync();
            _dialogDriver = async window =>
            {
                await Layout();
                check(window.GetVisualDescendants().OfType<Button>().Where(button => Equals(button.Content, UiText.Get("新建")) ||
                    Equals(button.Content, UiText.Get("重命名")) || Equals(button.Content, UiText.Get("上移"))).All(button => !button.IsEnabled),
                    "The sheet manager also disables mutations when opened from the sheet strip");
                window.Close();
            };
            await ShowSheetManagerAsync();
            await Layout();
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Bounds.Width, (int)Bounds.Height), new Vector(96, 96)))
            { bitmap.Render(this); bitmap.Save(Path.Combine(directory, "read-only.png"), PngBitmapEncoderOptions.Default); }
            await _workspace.ExportAsync(new(Path.Combine(directory, "read-only-export.xlsx")), false);
            check(Document.IsReadOnly && !Document.IsDirty && File.ReadAllBytes(source).SequenceEqual(original), "Export retains read-only state and original bytes");
            check(!await SaveAsync(false) && Document.IsReadOnly, "Canceling Save As leaves the source read-only");
            dialogs.NextSave = source;
            check(!await SaveAsync(false) && Document.IsReadOnly && File.ReadAllBytes(source).SequenceEqual(original), "Read-only Save refuses the original path");
            dialogs.NextSave = Path.Combine(directory, "read-only-copy.xlsq");
            check(await SaveAsync(false) && !Document.IsReadOnly && !_readOnlyNotice.IsVisible && !_formula.IsReadOnly && _addSheetButton.IsEnabled,
                "Save As succeeds to a writable copy and restores editing controls");
            check(await CommitCellAsync(0, new(0, 0), "editable copy") && Document.IsDirty && File.ReadAllBytes(source).SequenceEqual(original),
                "The new copy is editable while the source bytes remain unchanged");
            _workspace.New(); ShowHome();
            dialogs.NextOpen = dialogs.NextSave;
            Item(fileMenu, "以只读方式打开…").RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation;
            check(_book.IsVisible && Document.IsReadOnly && _workspace.FilePath == dialogs.NextSave,
                "Open Read-Only explicitly opens a writable file in read-only mode");
            _workspace.New(); ShowBook();
        }
        finally
        {
            _dialogDriver = null; _files = originalFiles; _clipboard = originalClipboard;
            if (OperatingSystem.IsWindows()) File.SetAttributes(source, FileAttributes.Normal);
        }
    }
}
