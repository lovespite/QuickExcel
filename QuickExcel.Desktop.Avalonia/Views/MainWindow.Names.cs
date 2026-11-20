using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Grid.Avalonia;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task NamesAsync()
    {
        if (await CommitEditorsAsync()) await ShowOwnedAsync<bool>(new NameManagerDialog(this));
    }

    private async Task CreateNameAsync()
    {
        if (await CommitEditorsAsync()) await EditNameAsync(this);
    }

    private async Task<DefinedName?> EditNameAsync(Window owner, DefinedName? selected = null)
    {
        if (!await CommitEditorsAsync()) return null;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId;
        var sheet = snapshot.GetSheet(snapshot.ActiveSheetId);
        var formula = selected?.Formula;
        if (formula is null)
        {
            var range = SelectedRange();
            formula = "=" + FormulaSyntax.Print(new ReferenceNode(sheet.Name, range.Top, range.Left, range.Bottom, range.Right,
                true, true, true, true, range.Count > 1));
        }
        var edit = new FieldsDialog(selected is null ? UiText.Get("新建名称") : UiText.Get("编辑名称"))
            .Text("name", UiText.Get("名称"), selected?.Name ?? "")
            .Choice("scope", UiText.Get("作用域"), new[] { UiText.Get("工作簿") }.Concat(snapshot.Sheets.Select(s => s.Name)).ToArray(),
                selected?.Scope is { } scope ? snapshot.Sheets.IndexOf(snapshot.GetSheet(scope)) + 1 : 0)
            .Text("comment", UiText.Get("注释"), selected?.Comment ?? "", multiline: true)
            .Text("formula", UiText.Get("引用位置 / 公式"), formula);
        if (selected is not null) edit.Input("scope").IsEnabled = false;
        DefinedName Definition() => new(edit.String("name").Trim(), edit.String("formula").Trim(),
            selected?.Scope ?? (edit.Index("scope") == 0 ? null : snapshot.Sheets[edit.Index("scope") - 1].Id), edit.String("comment"));
        edit.Validate = () =>
        {
            if (session != Document.Session.SessionId || snapshot.Revision != Document.Snapshot.Revision)
                throw new InvalidOperationException(UiText.Get("文档已改变，请关闭对话框后重试。"));
            var definition = Definition();
            if (snapshot.DefinedNames.Any(n => n != selected && n.Scope == definition.Scope && n.Name.Equals(definition.Name, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException(UiText.Get("此作用域内已存在同名名称。"));
            _ = snapshot.WithDefinedNames(snapshot.DefinedNames.Where(n => n != selected).Append(definition));
        };
        if (!await ShowOwnedAsync<bool>(edit, owner)) return null;
        if (session != Document.Session.SessionId) throw new InvalidOperationException(UiText.Get("文档已改变，请关闭对话框后重试。"));
        var result = Definition();
        DocumentCommand command = selected is null ? new DefineNameCommand(result) :
            new EditDefinedNameCommand(selected.Name, selected.Scope, result.Name, result.Formula, result.Comment);
        if (!await ExecuteAsync(command, snapshot.Revision)) return null;
        await CalculatePendingAsync();
        return Document.Snapshot.DefinedNames.Single(n => n.Scope == result.Scope && n.Name.Equals(result.Name, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class NameManagerDialog : Window
    {
        private readonly MainWindow _main;
        private readonly Guid _session;
        private readonly ListBox _list = new() { Name = "NameManagerList", SelectionMode = SelectionMode.Single };
        private readonly CompositionTextBox _search = new() { Name = "NameManagerSearch", PlaceholderText = UiText.Get("搜索名称、引用位置或注释") };
        private readonly ComboBox _filter = new() { Name = "NameManagerFilter", MinWidth = 145 };
        private readonly TextBox _reference = new() { Name = "NameManagerReference", IsReadOnly = true };
        private readonly TextBox _comment = new() { Name = "NameManagerComment", IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, Height = 62 };
        private readonly TextBlock _message = new() { Name = "NameManagerMessage", TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Firebrick };
        private readonly Button _create, _edit, _delete;
        private readonly SheetId[] _filterSheets;
        private bool _busy;
        public Task PendingOperation { get; private set; } = Task.CompletedTask;
        private WorkbookSnapshot Book => _main.Document.Snapshot;
        private DefinedName? Selected => (_list.SelectedItem as NameRow)?.Definition;
        private sealed record NameRow(DefinedName Definition, string Value, string Scope);
        private const string Columns = "1.1*,1.2*,1.8*,0.85*,1.2*";

        public NameManagerDialog(MainWindow main)
        {
            _main = main; _session = main.Document.Session.SessionId; _filterSheets = Book.Sheets.Select(s => s.Id).ToArray();
            Title = UiText.Get("名称管理器"); Width = 940; Height = 570; MinWidth = 720; MinHeight = 440;
            ShowInTaskbar = false; WindowStartupLocation = WindowStartupLocation.CenterOwner;
            var root = new LayoutGrid { Margin = new(16), RowDefinitions = new("Auto,10,Auto,8,Auto,*,12,Auto,6,Auto,8,Auto") };
            var toolbar = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            _create = ActionButton("NameManagerNew", UiText.Get("新建…"), async () => await _main.EditNameAsync(this));
            _edit = ActionButton("NameManagerEdit", UiText.Get("编辑…"), async () => await _main.EditNameAsync(this, Selected));
            _delete = ActionButton("NameManagerDelete", UiText.Get("删除"), async () =>
            {
                if (Selected is not { } selected) return null;
                if (await _main.ExecuteAsync(new DeleteDefinedNameCommand(selected.Name, selected.Scope), Book.Revision))
                    await _main.CalculatePendingAsync();
                return null;
            });
            toolbar.Children.Add(_create); toolbar.Children.Add(_edit); toolbar.Children.Add(_delete); root.Children.Add(toolbar);
            var filters = new LayoutGrid { ColumnDefinitions = new("*,10,Auto") };
            AutomationProperties.SetName(_search, UiText.Get("搜索名称、引用位置或注释"));
            AutomationProperties.SetName(_filter, UiText.Get("筛选作用域"));
            _filter.ItemsSource = new[] { UiText.Get("全部作用域"), UiText.Get("工作簿") }.Concat(Book.Sheets.Select(s => s.Name)).ToArray();
            _filter.SelectedIndex = 0;
            filters.Children.Add(_search); LayoutGrid.SetColumn(_filter, 2); filters.Children.Add(_filter); Put(filters, 2);
            var header = Row([UiText.Get("名称"), UiText.Get("数值"), UiText.Get("引用位置"), UiText.Get("作用域"), UiText.Get("注释")], true);
            Put(new Border { Background = new SolidColorBrush(Color.Parse("#F0F3F7")), Padding = new(7, 6), Child = header }, 4);
            _list.ItemTemplate = new FuncDataTemplate<NameRow>((item, _) => item is null ? null : Row([item.Definition.Name, item.Value, item.Definition.Formula, item.Scope, item.Definition.Comment], false));
            _list.Styles.Add(new Style(s => s.OfType<ListBoxItem>()) { Setters =
                { new Setter(PaddingProperty, new Thickness(7, 6)), new Setter(HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch) } });
            AutomationProperties.SetName(_list, UiText.Get("已定义名称"));
            Put(_list, 5);
            var reference = new LayoutGrid { ColumnDefinitions = new("80,*") };
            reference.Children.Add(new TextBlock { Text = UiText.Get("引用位置"), VerticalAlignment = VerticalAlignment.Center });
            LayoutGrid.SetColumn(_reference, 1); reference.Children.Add(_reference); Put(reference, 7);
            var comments = new LayoutGrid { ColumnDefinitions = new("80,*") };
            comments.Children.Add(new TextBlock { Text = UiText.Get("注释"), VerticalAlignment = VerticalAlignment.Top, Margin = new(0, 6, 0, 0) });
            LayoutGrid.SetColumn(_comment, 1); comments.Children.Add(_comment); Put(comments, 9);
            var footer = new DockPanel();
            var close = new Button { Name = "NameManagerClose", Content = UiText.Get("关闭"), IsCancel = true, MinWidth = 85 };
            close.Click += (_, _) => Close(); DockPanel.SetDock(close, Dock.Right); footer.Children.Add(close); footer.Children.Add(_message); Put(footer, 11);
            Content = root;
            _list.SelectionChanged += (_, _) => UpdateSelection();
            _list.DoubleTapped += (_, _) => { if (_edit.IsEnabled) PendingOperation = Run(async () => await _main.EditNameAsync(this, Selected)); };
            _search.TextChanged += (_, _) => Reload(Selected);
            _filter.SelectionChanged += (_, _) => Reload(Selected);
            Reload(); Opened += (_, _) => _search.Focus();
            void Put(Control control, int row) { LayoutGrid.SetRow(control, row); root.Children.Add(control); }
        }

        private static Control Row(string[] values, bool header)
        {
            var grid = new LayoutGrid { ColumnDefinitions = new(Columns) };
            for (var i = 0; i < values.Length; i++)
            {
                var text = new TextBlock { Text = values[i], TextTrimming = TextTrimming.CharacterEllipsis,
                    FontWeight = header ? FontWeight.SemiBold : FontWeight.Normal, Margin = new(0, 0, 10, 0) };
                ToolTip.SetTip(text, values[i]); LayoutGrid.SetColumn(text, i); grid.Children.Add(text);
            }
            return grid;
        }

        private Button ActionButton(string name, string text, Func<Task<DefinedName?>> action)
        {
            var button = new Button { Name = name, Content = text, MinWidth = 90 };
            button.Click += (_, _) => PendingOperation = Run(action);
            return button;
        }

        private async Task Run(Func<Task<DefinedName?>> action)
        {
            if (_busy || _main.Document.IsReadOnly) return;
            if (_session != _main.Document.Session.SessionId) { _message.Text = UiText.Get("文档已改变，请关闭对话框后重试。"); return; }
            var selected = Selected; _busy = true; UpdateSelection(); _message.Text = "";
            try { selected = await action() ?? selected; }
            catch (Exception ex) { _message.Text = ex.Message; }
            finally { _busy = false; Reload(selected); }
        }

        private void Reload(DefinedName? selected = null)
        {
            var text = _search.Text?.Trim() ?? "";
            var sheet = _filter.SelectedIndex > 1 && _filter.SelectedIndex - 2 < _filterSheets.Length ? _filterSheets[_filter.SelectedIndex - 2] : (SheetId?)null;
            var rows = Book.DefinedNames.Where(n => (_filter.SelectedIndex <= 0 || n.Scope == sheet) &&
                    (text.Length == 0 || n.Name.Contains(text, StringComparison.OrdinalIgnoreCase) || n.Formula.Contains(text, StringComparison.OrdinalIgnoreCase) || n.Comment.Contains(text, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(n => n.Name, StringComparer.OrdinalIgnoreCase).ThenBy(n => n.Scope is null ? "" : Book.GetSheet(n.Scope.Value).Name, StringComparer.OrdinalIgnoreCase)
                .Select(n => new NameRow(n, NamePreview(Book, n), n.Scope is { } id ? Book.GetSheet(id).Name : UiText.Get("工作簿"))).ToArray();
            _list.ItemsSource = rows;
            _list.SelectedItem = rows.FirstOrDefault(r => r.Definition.Scope == selected?.Scope && r.Definition.Name.Equals(selected?.Name, StringComparison.OrdinalIgnoreCase)) ?? rows.FirstOrDefault();
            UpdateSelection();
        }

        private void UpdateSelection()
        {
            var writable = !_busy && _session == _main.Document.Session.SessionId && !_main.Document.IsReadOnly && Book.FileState?.FormulaReadOnly != true;
            _create.IsEnabled = writable; _edit.IsEnabled = _delete.IsEnabled = writable && Selected is not null;
            _list.IsEnabled = _search.IsEnabled = _filter.IsEnabled = !_busy;
            _reference.Text = Selected?.Formula ?? ""; _comment.Text = Selected?.Comment ?? "";
        }
    }

    private static string NamePreview(WorkbookSnapshot book, DefinedName definition)
    {
        try
        {
            var expression = FormulaBinding.Bind(FormulaSyntax.Parse(definition.Formula), book, definition.Scope ?? book.ActiveSheetId, arrayContext: true);
            if (expression is ValueNode literal) return CellDisplayFormatter.Format(literal.Value, book.Styles[0], book.DateSystem);
            if (expression is not ReferenceNode reference) return UiText.Get("公式");
            var sheet = book.Sheets.Single(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
            var values = new List<string>();
            var rows = reference.EndRow - reference.Row + 1; var columns = reference.EndColumn - reference.Column + 1;
            for (var i = 0; i < Math.Min(6L, (long)rows * columns); i++)
            {
                var row = reference.Row + i / columns; var column = reference.Column + i % columns;
                var cell = sheet.GetCell(row, column);
                var value = cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null) && cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error)
                    ? UiText.Get("待计算") : CellDisplayFormatter.Format(cell?.Value ?? CellValue.Blank, book.Styles[sheet.GetStyleId(row, column)], book.DateSystem);
                values.Add(value.Length > 80 ? value[..80] + "…" : value);
            }
            if ((long)rows * columns == 1) return values[0];
            return "{" + string.Join(", ", values) + ((long)rows * columns > 6 ? ", …" : "") + "}";
        }
        catch (Exception ex) when (ex is FormatException or ArgumentException or InvalidOperationException or NotSupportedException)
        { return UiText.Get("无法预览"); }
    }
}
