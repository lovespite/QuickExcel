using System.Collections.Immutable;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class TextSplitWizard : Window
{
    private readonly WorkbookSnapshot _book;
    private readonly CellRectangle _source;
    private readonly string[] _rows;
    private readonly CleaningColumnType[] _formats = Enumerable.Repeat(CleaningColumnType.General, 256).ToArray();
    private readonly TextSplitPreview _preview = new();
    private readonly TextBlock _instruction = Label("");
    private readonly TextBlock _caption = Label("");
    private readonly TextBlock _error = new() { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap };
    private readonly TextBox _destination = new();
    private readonly CheckBox _overwrite = new() { IsVisible = false };
    private readonly RadioButton _fixed = Radio("固定宽度", "mode");
    private readonly CheckBox[] _delimiters = new[] { "Tab 键", "分号", "逗号", "空格", "其他" }.Select(s => new CheckBox { Content = s }).ToArray();
    private readonly TextBox _other = new() { Width = 68 };
    private readonly TextBox _breaks = new() { PlaceholderText = "例如：3,7" };
    private readonly CheckBox _collapse = new() { Content = "连续分隔符视为单个处理" };
    private readonly ComboBox _qualifier = Choice(["双引号（\"）", "单引号（'）", "无"]);
    private readonly ComboBox _dateOrder = Choice(["YMD", "MDY", "DMY"]);
    private static readonly CleaningColumnType[] DateTypes = [CleaningColumnType.DateYmd, CleaningColumnType.DateMdy, CleaningColumnType.DateDmy];
    private readonly Dictionary<CleaningColumnType, RadioButton> _formatButtons = [];
    private readonly Button _back = new() { Content = "上一步", MinWidth = 88 };
    private readonly Button _next = new() { Content = "下一步", MinWidth = 88 };
    private readonly Button _finish = new() { Content = "完成", MinWidth = 88 };
    private readonly StackPanel _options = new() { MinHeight = 216 };
    private readonly Control _modePanel, _delimiterPanel, _breakPanel, _formatPanel;
    private NumericTextOptions _numbers = new();
    private CleaningPreview? _candidate;
    private int _step = 1, _column;
    private bool _syncing, _busy, _closed, _sampleValid;

    public CleaningPreview? Result { get; private set; }
    public bool AllowOverwrite => _overwrite.IsChecked == true;

    public TextSplitWizard(WorkbookSnapshot book, CellRectangle source, string[] rows)
    {
        _book = book;
        _source = source;
        _rows = rows;
        Width = 680; Height = 610; MinWidth = 600; MinHeight = 530;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ShowInTaskbar = false;
        _destination.Text = $"${FormulaSyntax.ColumnName(source.Left)}${source.Top + 1}";
        AutomationProperties.SetName(_destination, "目标区域");
        _modePanel = BuildMode();
        _delimiterPanel = BuildDelimiters();
        _breakPanel = Group("固定宽度", Label("输入每个分界前的累计字符宽度，用逗号分隔。"), _breaks);
        _formatPanel = BuildFormats();
        foreach (var page in new[] { _modePanel, _delimiterPanel, _breakPanel, _formatPanel }) _options.Children.Add(page);

        var layout = new LayoutGrid { Margin = new(16), RowDefinitions = new("Auto,Auto,*,Auto,Auto"), RowSpacing = 12 };
        layout.Children.Add(_instruction);
        Add(layout, _options, 1);
        var table = new LayoutGrid { Margin = new(0, 12, 0, 0), RowDefinitions = new("Auto,*"), RowSpacing = 6 };
        table.Children.Add(_caption); Add(table, _preview, 1); Add(layout, table, 2);
        var messages = new StackPanel();
        messages.Children.Add(_error); messages.Children.Add(_overwrite); Add(layout, messages, 3);
        var cancel = new Button { Content = "取消", MinWidth = 88, IsCancel = true };
        var buttons = Row(cancel, _back, _next, _finish);
        buttons.HorizontalAlignment = HorizontalAlignment.Right;
        Add(layout, buttons, 4);
        Content = layout;

        cancel.Click += (_, _) => Close(false);
        Closed += (_, _) => _closed = true;
        _back.Click += (_, _) => ShowStep(_step - 1);
        _next.Click += async (_, _) => await AdvanceAsync(false);
        _finish.Click += async (_, _) => await AdvanceAsync(true);
        _preview.ColumnSelected += column => { _column = column; SyncFormat(); };
        _destination.TextChanging += (_, _) => InvalidateCandidate();
        _overwrite.IsCheckedChanged += (_, _) => UpdateButtons();
        foreach (var box in _delimiters.Append(_collapse)) box.IsCheckedChanged += (_, _) => SplitChanged();
        _other.TextChanging += (_, _) => SplitChanged();
        _breaks.TextChanging += (_, _) => SplitChanged();
        _qualifier.SelectionChanged += (_, _) => SplitChanged();
        _fixed.IsCheckedChanged += (_, _) => SplitChanged();
        ShowStep(1);
    }

    private static TextBlock Label(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };
    private static RadioButton Radio(string label, string group) => new() { Content = label, GroupName = group, MinHeight = 28 };
    private static ComboBox Choice(string[] values) => new() { ItemsSource = values, SelectedIndex = 0, MinWidth = 86 };
    private static void Add(LayoutGrid grid, Control child, int row) { LayoutGrid.SetRow(child, row); grid.Children.Add(child); }
    private static StackPanel Row(params Control[] children)
    {
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };
        foreach (var child in children) row.Children.Add(child);
        return row;
    }
    private static Border Group(string title, params Control[] children)
    {
        var stack = new StackPanel { Spacing = 3 };
        stack.Children.Add(new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, Margin = new(0, 0, 0, 4) });
        foreach (var child in children) stack.Children.Add(child);
        return new Border { Child = stack, Padding = new(12, 8), BorderThickness = new(1), BorderBrush = Brushes.LightGray };
    }

    private Control BuildMode()
    {
        var delimited = Radio("分隔符号", "mode"); delimited.IsChecked = true;
        return Group("原始数据类型", Row(delimited, Label("用逗号、制表符等字符分隔字段")),
            Row(_fixed, Label("按指定的字符宽度切分字段")));
    }
    private Control BuildDelimiters()
    {
        string chars = "\t;, ";
        var sample = _rows.Take(TextSplitPreview.SampleLimit).ToArray();
        var initial = Array.FindIndex(chars.ToCharArray(), c => c != ' ' && sample.Any(s => s.Contains(c)));
        _delimiters[Math.Max(initial, 0)].IsChecked = true;
        AutomationProperties.SetName(_other, "其他分隔符");
        AutomationProperties.SetName(_breaks, "固定宽度分界");
        AutomationProperties.SetName(_qualifier, "文本识别符号");
        var choices = Group("分隔符号", _delimiters[0], _delimiters[1], _delimiters[2], _delimiters[3], Row(_delimiters[4], _other));
        var options = new StackPanel { Spacing = 12, VerticalAlignment = VerticalAlignment.Center };
        options.Children.Add(_collapse); options.Children.Add(Row(Label("文本识别符号"), _qualifier));
        return Row(choices, options);
    }
    private Control BuildFormats()
    {
        RadioButton Format(CleaningColumnType type, string name)
        {
            var button = Radio(name, "format");
            _formatButtons.Add(type, button);
            button.IsCheckedChanged += (_, _) =>
            {
                if (_syncing || button.IsChecked != true) return;
                SetFormat(type == CleaningColumnType.DateYmd ? DateTypes[_dateOrder.SelectedIndex] : type);
            };
            return button;
        }
        var date = Format(CleaningColumnType.DateYmd, "日期");
        AutomationProperties.SetName(_dateOrder, "日期顺序");
        _dateOrder.SelectionChanged += (_, _) =>
        {
            if (!_syncing && _dateOrder.SelectedIndex >= 0) SetFormat(DateTypes[_dateOrder.SelectedIndex]);
        };
        var formats = Group("列数据格式", Format(CleaningColumnType.General, "常规"), Format(CleaningColumnType.Text, "文本"),
            Row(date, _dateOrder), Format(CleaningColumnType.Skip, "不导入此列（跳过）"), Format(CleaningColumnType.Number, "数字（严格转换）"));
        var advanced = new Button { Content = "高级…" };
        advanced.Click += async (_, _) =>
        {
            var dialog = MainWindow.NumericFields(new FieldsDialog("高级分列设置"), _numbers);
            dialog.Validate = () => MainWindow.ReadNumeric(dialog);
            if (await dialog.ShowAsync(this)) { _numbers = MainWindow.ReadNumeric(dialog); InvalidateCandidate(); }
        };
        var help = new StackPanel { MaxWidth = 315, Spacing = 12, VerticalAlignment = VerticalAlignment.Center };
        help.Children.Add(Label("“常规”会识别数值和年在前的日期，其余内容保留为文本。要保留编号的前导零，请选“文本”。"));
        help.Children.Add(advanced);
        var target = new DockPanel();
        var label = Label("目标区域  "); DockPanel.SetDock(label, Dock.Left); target.Children.Add(label); target.Children.Add(_destination);
        var panel = new StackPanel { Spacing = 10 };
        panel.Children.Add(Row(formats, help)); panel.Children.Add(target);
        return panel;
    }

    private void SetFormat(CleaningColumnType type)
    {
        _formats[_column] = type;
        InvalidateCandidate();
        SyncFormat();
    }
    private void SyncFormat()
    {
        _syncing = true;
        try
        {
            var type = _formats[_column];
            var date = Array.IndexOf(DateTypes, type);
            if (date >= 0) _dateOrder.SelectedIndex = date;
            _formatButtons[date >= 0 ? CleaningColumnType.DateYmd : type].IsChecked = true;
            _preview.SetFormats(_column, i => _formats[i] switch
            {
                CleaningColumnType.Text => "文本", CleaningColumnType.Number => "数字", CleaningColumnType.Skip => "跳过",
                CleaningColumnType.DateYmd => "日期 YMD", CleaningColumnType.DateMdy => "日期 MDY", CleaningColumnType.DateDmy => "日期 DMY", _ => "常规"
            });
        }
        finally { _syncing = false; }
    }

    private TextSplitOptions ReadOptions()
    {
        if (_fixed.IsChecked == true)
        {
            var breaks = (_breaks.Text ?? "").Split([',', '，'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            return new(Breaks: breaks.Select(s => int.TryParse(s, out var n) ? n : throw new ArgumentException("固定宽度分界必须是整数。")).ToImmutableArray());
        }
        string[] chars = ["\t", ";", ",", " ", _other.Text ?? ""];
        var delimiters = string.Concat(chars.Where((_, i) => _delimiters[i].IsChecked == true));
        return new(delimiters, _collapse.IsChecked == true, _qualifier.SelectedIndex switch { 0 => '"', 1 => '\'', _ => null });
    }
    private void SplitChanged()
    {
        if (_closed) return;
        Array.Fill(_formats, CleaningColumnType.General);
        _column = 0;
        InvalidateCandidate();
        if (_step == 2) ShowStep(2);
    }
    private void InvalidateCandidate()
    {
        _candidate = null;
        _overwrite.IsVisible = false;
        _overwrite.IsChecked = false;
        _error.Text = "";
        UpdateButtons();
    }

    private void UpdateButtons()
    {
        _back.IsEnabled = !_busy && _step > 1;
        _next.IsEnabled = !_busy && _step < 3 && _sampleValid;
        _finish.IsEnabled = !_busy && _sampleValid && (!_overwrite.IsVisible || AllowOverwrite);
    }

    private void ShowStep(int step, int fullWidth = 0)
    {
        _step = step;
        InvalidateCandidate();
        Title = $"文本分列向导 - 第 {step} 步，共 3 步";
        _instruction.Text = step switch { 1 => "请选择最适合选定数据的分列方式。", 2 when _fixed.IsChecked == true => "请设置分界位置，并在下方预览切分结果。",
            2 => "请选择数据包含的分隔符号，并在下方预览分列效果。", _ => "请在预览中单击一列，然后为该列设置数据格式。" };
        _modePanel.IsVisible = step == 1;
        _delimiterPanel.IsVisible = step == 2 && _fixed.IsChecked != true;
        _breakPanel.IsVisible = step == 2 && _fixed.IsChecked == true;
        _formatPanel.IsVisible = step == 3;
        _other.IsEnabled = _delimiters[4].IsChecked == true;
        _next.IsDefault = step < 3; _finish.IsDefault = step == 3;
        _caption.Text = step == 1 ? "预览选定数据" : "数据预览";
        if (_rows.Length > TextSplitPreview.SampleLimit) _caption.Text += $"（显示前 {TextSplitPreview.SampleLimit} 行，共 {_rows.Length:N0} 行）";
        var sample = _rows.Take(TextSplitPreview.SampleLimit);
        try
        {
            var split = step == 1 ? null : ReadOptions();
            var parsed = sample.Select((row, i) => split is null ? [row] : split.SplitRow(row, _source.Top + i)).ToArray();
            var width = Math.Max(fullWidth, parsed.Max(row => row.Length));
            _preview.ShowRows(parsed, width, step == 1 ? _source.Top : null, step == 3);
            _sampleValid = true;
            if (step == 3) { _column = Math.Min(_column, width - 1); SyncFormat(); }
        }
        catch (ArgumentException ex)
        {
            _preview.ShowRows([], 0);
            _sampleValid = false;
            _error.Text = ex.Message;
        }
        UpdateButtons();
    }

    private Task<int> CountColumnsAsync(TextSplitOptions options) =>
        Task.Run(() => _rows.Select((row, i) => options.SplitRow(row, _source.Top + i).Length).Max());

    private async Task AdvanceAsync(bool finish)
    {
        if (_busy || _closed) return;
        _busy = true;
        _options.IsEnabled = _preview.IsEnabled = false;
        UpdateButtons();
        try
        {
            if (!finish && _step == 1) { ShowStep(2); return; }
            var options = ReadOptions();
            if (!finish)
            {
                var width = await CountColumnsAsync(options);
                if (!_closed) ShowStep(3, width);
                return;
            }
            var target = CellRectangle.Parse((_destination.Text ?? "").Trim());
            if (target.Count != 1) throw new ArgumentException("目标区域必须是单个起始单元格。");
            var formats = _formats.ToArray();
            var result = _candidate ?? await Task.Run(() => DataCleaning.SplitText(_book, _book.ActiveSheetId, _source,
                new(target.Top, target.Left), options, formats, _numbers));
            if (_closed) return;
            if (!result.Overwrites.IsEmpty && !AllowOverwrite)
            {
                if (_step != 3)
                {
                    var width = await CountColumnsAsync(options);
                    if (_closed) return;
                    ShowStep(3, width);
                }
                _candidate = result;
                var examples = result.Overwrites.Take(4).Select(a => FormulaSyntax.ColumnName(a.Column) + (a.Row + 1));
                _overwrite.Content = $"确认覆盖 {result.Overwrites.Length:N0} 个已有内容的目标单元格（{string.Join("、", examples)}）";
                _overwrite.IsVisible = true;
                return;
            }
            Result = result;
            Close(true);
        }
        catch (Exception ex) { if (!_closed) _error.Text = ex.Message; }
        finally
        {
            _busy = false;
            _options.IsEnabled = _preview.IsEnabled = true;
            UpdateButtons();
        }
    }
}
