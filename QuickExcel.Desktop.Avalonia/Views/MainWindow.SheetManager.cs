using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task ShowSheetManagerAsync()
    {
        if (await CommitEditorsAsync()) await ShowOwnedAsync<bool>(new SheetManagerDialog(this));
    }

    private sealed class SheetManagerDialog : Window
    {
        private readonly MainWindow _main;
        private readonly ListBox _list = new() { SelectionMode = SelectionMode.Single };
        private readonly CompositionTextBox _search = new() { PlaceholderText = UiText.Get("搜索工作表") };
        private readonly StackPanel _actions = new() { Spacing = 5 };
        private readonly TextBlock _message = new() { TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };
        private readonly List<(Control Control, Func<bool> Allowed)> _availability = [];
        private readonly Button _visibility;
        private SheetId? _selectAfter;
        private bool _busy;
        private WorkbookSnapshot Book => _main.Document.Snapshot;
        private WorksheetSnapshot? Selected => _list.SelectedItem as WorksheetSnapshot;

        public SheetManagerDialog(MainWindow main)
        {
            _main = main;
            Title = UiText.Get("工作表管理器"); Width = 580; Height = 490; MinWidth = 420; MinHeight = 360;
            ShowInTaskbar = false; WindowStartupLocation = WindowStartupLocation.CenterOwner;
            var root = new LayoutGrid { Margin = new(16), ColumnDefinitions = new("*,16,108"), RowDefinitions = new("Auto,8,*,12,Auto") };
            var searchBar = new DockPanel();
            var find = new Button { Content = UiText.Get("查找下一个"), Margin = new(6, 0, 0, 0) };
            find.Click += (_, _) => Find(next: true);
            _search.TextChanged += (_, _) => Find(next: false);
            _search.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter && !_search.IsComposing) { e.Handled = true; Find(next: true); }
            };
            AutomationProperties.SetName(_search, UiText.Get("搜索工作表"));
            DockPanel.SetDock(find, Dock.Right); searchBar.Children.Add(find); searchBar.Children.Add(_search);
            root.Children.Add(searchBar);
            _list.ItemTemplate = new FuncDataTemplate<WorksheetSnapshot>((sheet, _) => sheet is null ? null : SheetRow(sheet));
            AutomationProperties.SetName(_list, UiText.Get("所有工作表"));
            LayoutGrid.SetRow(_list, 2); root.Children.Add(_list);
            var buttons = new ScrollViewer { Content = _actions };
            LayoutGrid.SetColumn(buttons, 2); LayoutGrid.SetRowSpan(buttons, 3); root.Children.Add(buttons);

            ActionButton(UiText.Get("上移"), () => Move(-1), () => _list.SelectedIndex > 0);
            ActionButton(UiText.Get("下移"), () => Move(1), () => Selected is not null && _list.SelectedIndex < Book.Sheets.Length - 1);
            ActionButton(UiText.Get("新建"), async () => { await main.AddSheetAsync(); _selectAfter = Book.ActiveSheetId; }, HasCapacity);
            ActionButton(UiText.Get("复制"), async () => _selectAfter = await main.CopySheetAsync(this, Selected!.Id) ?? _selectAfter,
                () => Selected is not null && HasCapacity());
            ActionButton(UiText.Get("重命名"), () => main.RenameSheetAsync(this, Selected!.Id));
            ActionButton(UiText.Get("保护…"), () => main.ProtectSheetAsync(this, Selected!.Id));
            _visibility = ActionButton(UiText.Get("隐藏"), () => main.ExecuteAsync(new SetSheetVisibilityCommand(Selected!.Id,
                Selected.Visibility == SheetVisibility.Visible ? SheetVisibility.Hidden : SheetVisibility.Visible)),
                () => Selected?.Visibility == SheetVisibility.Hidden || CanRemoveVisible());
            var colors = new DropDownButton { Content = UiText.Get("设置颜色"), HorizontalAlignment = HorizontalAlignment.Stretch };
            MenuItem ColorItem(string title, Func<Task> action)
            {
                var item = new MenuItem { Header = title };
                item.Click += async (_, _) => await Run(action);
                return item;
            }
            colors.Flyout = new MenuFlyout { ItemsSource = main.ColorMenuItems(UiText.Get("恢复默认颜色"),
                color => main.ExecuteAsync(new SetSheetTabColorCommand(Selected!.Id, color)),
                () => main.CustomSheetTabColorAsync(this, Selected!.Id), ColorItem) };
            _availability.Add((colors, () => Selected is not null)); _actions.Children.Add(colors);
            ActionButton(UiText.Get("删除…"), () => main.DeleteSheetAsync(this, Selected!.Id),
                () => Selected is { Visibility: not SheetVisibility.Visible } || CanRemoveVisible());

            var footer = new DockPanel();
            var close = new Button { Content = UiText.Get("关闭"), IsCancel = true, MinWidth = 78, HorizontalContentAlignment = HorizontalAlignment.Center };
            close.Click += (_, _) => Close();
            DockPanel.SetDock(close, Dock.Right); footer.Children.Add(close); footer.Children.Add(_message);
            LayoutGrid.SetRow(footer, 4); LayoutGrid.SetColumnSpan(footer, 3); root.Children.Add(footer);
            Content = root;
            _list.SelectionChanged += (_, _) => UpdateActions();
            Reload(Book.ActiveSheetId, 0);
            Opened += (_, _) => _search.Focus();
        }

        private static Control SheetRow(WorksheetSnapshot sheet)
        {
            var row = new DockPanel { LastChildFill = true };
            var flags = new List<string>();
            if (sheet.Visibility == SheetVisibility.Hidden) flags.Add(UiText.Get("隐藏"));
            if (sheet.Visibility == SheetVisibility.VeryHidden) flags.Add(UiText.Get("深度隐藏"));
            if (sheet.Data.Protection?.Sheet == true) flags.Add(UiText.Get("已保护"));
            var state = string.Join(" · ", flags);
            var status = new TextBlock { Text = state, FontSize = 12, Opacity = 0.75, Margin = new(8, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
            DockPanel.SetDock(status, Dock.Right); row.Children.Add(status);
            var swatch = new Border { Width = 5, CornerRadius = new(2), Margin = new(0, 2, 9, 2),
                Background = ColorPalette.TryParse(sheet.TabColor, out var color) ? new SolidColorBrush(color) : Brushes.Transparent };
            DockPanel.SetDock(swatch, Dock.Left); row.Children.Add(swatch);
            row.Children.Add(new TextBlock { Text = sheet.Name, TextTrimming = TextTrimming.CharacterEllipsis });
            ToolTip.SetTip(row, sheet.Name); AutomationProperties.SetName(row, sheet.Name + " " + state);
            return row;
        }

        private bool HasCapacity() => Book.Sheets.Length < WorkbookLimits.MaxSheets;
        private bool CanRemoveVisible() => Selected?.Visibility == SheetVisibility.Visible && Book.Sheets.Count(s => s.Visibility == SheetVisibility.Visible) > 1;

        private Button ActionButton(string text, Func<Task> run, Func<bool>? allowed = null)
        {
            var button = new Button { Content = text, HorizontalAlignment = HorizontalAlignment.Stretch };
            button.Click += async (_, _) => await Run(run);
            _actions.Children.Add(button); _availability.Add((button, allowed ?? (() => Selected is not null))); return button;
        }

        private Task Move(int direction) => _main.ExecuteAsync(new ReorderSheetCommand(Selected!.Id, _list.SelectedIndex + direction), Book.Revision);

        private void UpdateActions()
        {
            foreach (var (control, allowed) in _availability)
                control.IsEnabled = !_busy && !_main.Document.IsReadOnly && Book.FileState?.FormulaReadOnly != true && allowed();
            _visibility.Content = Selected?.Visibility == SheetVisibility.Hidden ? UiText.Get("取消隐藏") : UiText.Get("隐藏");
        }

        private async Task Run(Func<Task> action)
        {
            if (_busy || _main.Document.IsReadOnly) return;
            var index = _list.SelectedIndex; _selectAfter = Selected?.Id;
            _busy = true; UpdateActions(); _message.Text = "";
            try { await action(); }
            catch (Exception ex) { _message.Text = ex.Message; }
            finally { _busy = false; Reload(_selectAfter, index); }
        }

        private void Reload(SheetId? id, int fallback)
        {
            _list.ItemsSource = Book.Sheets;
            var sheet = Book.Sheets.FirstOrDefault(s => s.Id == id);
            _list.SelectedIndex = sheet is null ? Math.Min(fallback, Book.Sheets.Length - 1) : Book.Sheets.IndexOf(sheet);
            if (_list.SelectedIndex >= 0) _list.ScrollIntoView(_list.SelectedIndex);
            UpdateActions();
        }

        private void Find(bool next)
        {
            var query = _search.Text ?? "";
            var start = next ? _list.SelectedIndex + 1 : 0;
            for (var i = 0; query.Length > 0 && i < _list.ItemCount; i++)
            {
                var index = (start + i) % _list.ItemCount;
                if (_list.Items[index] is WorksheetSnapshot sheet && sheet.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    _list.SelectedIndex = index; _list.ScrollIntoView(index); _message.Text = ""; return;
                }
            }
            _list.SelectedIndex = -1;
            _message.Text = query.Length == 0 ? "" : UiText.Get("没有匹配的工作表");
        }
    }
}
