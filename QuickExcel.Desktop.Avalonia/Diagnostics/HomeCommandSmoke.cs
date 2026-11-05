using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckHomeCommandAvailabilityAsync(string path, Action<bool, string> check)
    {
        var fileMenu = (MenuItem)_menu.Items[0]!;
        MenuItem FileItem(string title) => fileMenu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get(title)));
        async Task Click(MenuItem item) { item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation; }
        async Task Key(Key key, KeyModifiers modifiers = KeyModifiers.Control)
        {
            RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = key, KeyModifiers = modifiers });
            await _operation;
        }
        void CheckHome()
        {
            check(_start.IsVisible && !_book.IsVisible && _menu.Items.OfType<MenuItem>().Skip(1).All(item => !item.IsEnabled),
                "Home disables Edit, Worksheet, View and Script menus");
            check(new[] { "保存  Ctrl+S", "另存为", "导出值快照", "导出可移植公式", "打印预览 / PDF", "兼容性报告", "关闭工作簿" }
                .All(title => !FileItem(title).IsEnabled), "Home disables workbook-dependent File commands");
            check(new[] { "新建  Ctrl+N", "打开  Ctrl+O", "最近文件", "打开恢复副本", "设置", "退出" }.All(title => FileItem(title).IsEnabled),
                "Home retains New, Open, Recent, Recovery, Settings and Exit");
        }
        var originalFiles = _files; var originalChoice = _discardChoice;
        try
        {
            _files = new SmokeFiles { NextOpen = path };
            _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            CheckHome();
            var document = Document; var snapshot = document.Snapshot;
            foreach (var menu in _menu.Items.OfType<MenuItem>().Skip(1))
                foreach (var item in menu.Items.OfType<MenuItem>()) await Click(item);
            foreach (var title in new[] { "保存  Ctrl+S", "另存为", "导出值快照", "导出可移植公式", "打印预览 / PDF", "兼容性报告", "关闭工作簿" })
                await Click(FileItem(title));
            foreach (var key in new[] { global::Avalonia.Input.Key.S, global::Avalonia.Input.Key.Z, global::Avalonia.Input.Key.Y,
                global::Avalonia.Input.Key.V, global::Avalonia.Input.Key.X, global::Avalonia.Input.Key.D, global::Avalonia.Input.Key.R,
                global::Avalonia.Input.Key.E, global::Avalonia.Input.Key.F, global::Avalonia.Input.Key.G, global::Avalonia.Input.Key.W }) await Key(key);
            await Key(global::Avalonia.Input.Key.Delete, KeyModifiers.None);
            await Key(global::Avalonia.Input.Key.F9, KeyModifiers.None);
            await Key(global::Avalonia.Input.Key.F6, KeyModifiers.None);
            check(ReferenceEquals(Document, document) && ReferenceEquals(Document.Snapshot, snapshot) && !Document.IsDirty && !_book.IsVisible,
                "Routed disabled menu clicks and workbook shortcuts cannot mutate the hidden startup document");
            _recentSearch.Focus(); _recentSearch.Text = "home search";
            var delete = new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = global::Avalonia.Input.Key.Delete };
            Shortcuts(this, delete);
            check(!delete.Handled && !Document.IsDirty, "Home leaves text-editing keys to the recent-file search");
            _recentSearch.Text = "";

            await Key(global::Avalonia.Input.Key.N);
            check(_book.IsVisible && _workbookMenuItems.All(item => item.IsEnabled) && _menu.Items.OfType<MenuItem>().All(item => item.IsEnabled),
                "Ctrl+N enables workbook menus after creating a workbook");
            await Click(((MenuItem)_menu.Items[2]!).Items.OfType<MenuItem>().Single(item => Equals(item.Header, UiText.Get("添加工作表"))));
            check(Document.Snapshot.Sheets.Length == 2 && Document.IsDirty, "Worksheet menu works after New");
            _discardChoice = () => Task.FromResult(SaveDecision.Cancel);
            await Key(global::Avalonia.Input.Key.W);
            check(_book.IsVisible && FileItem("保存  Ctrl+S").IsEnabled && Document.IsDirty, "Canceling Close retains the enabled workbook");
            _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            await Key(global::Avalonia.Input.Key.W); CheckHome();
            await Key(global::Avalonia.Input.Key.O);
            check(_book.IsVisible && _workspace.FilePath == path && FileItem("保存  Ctrl+S").IsEnabled, "Ctrl+O opens a workbook and restores its commands");
            try { using var external = File.OpenRead(path); throw new InvalidOperationException("Opened XLSQ was not locked"); }
            catch (IOException) { check(true, "The native Open command retains exclusive XLSQ access"); }
            await Click(FileItem("关闭工作簿")); CheckHome();
            using var released = File.OpenRead(path);
            check(released.Length > 0, "Close Workbook releases the XLSQ lock before returning home");
        }
        finally { _files = originalFiles; _discardChoice = originalChoice; _recentSearch.Text = ""; }
    }
}
