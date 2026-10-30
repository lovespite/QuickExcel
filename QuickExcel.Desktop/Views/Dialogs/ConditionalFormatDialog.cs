using System.Windows;
using System.Windows.Controls;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed class ConditionalRuleDialog : Window
{
    internal ComboBox KindPicker { get; } = new() { ItemsSource = new[] { "单元格比较", "公式", "重复值", "唯一值", "包含文本", "不包含文本", "以文本开头", "以文本结尾" }, SelectedIndex = 0 };
    internal ComboBox OperatorPicker { get; } = new() { ItemsSource = new[] { "介于", "不介于", "等于", "不等于", "小于", "小于或等于", "大于", "大于或等于" }, SelectedIndex = 0 };
    internal TextBox Ranges { get; } = new();
    internal TextBox First { get; } = new() { Text = "1" };
    internal TextBox Second { get; } = new() { Text = "10" };
    internal TextBox MatchText { get; } = new();
    internal TextBox Fill { get; } = new() { Text = "FFFFC7CE" };
    internal TextBox Font { get; } = new() { Text = "FF9C0006" };
    internal TextBox NumberFormat { get; } = new();
    internal CheckBox Bold { get; } = new() { Content = "粗体（灰色：不改变）", IsThreeState = true, IsChecked = null };
    internal CheckBox Italic { get; } = new() { Content = "斜体（灰色：不改变）", IsThreeState = true, IsChecked = null };
    internal CheckBox Stop { get; } = new() { Content = "为真时停止后续规则" };
    internal Button ApplyButton { get; } = new() { Content = "确定", IsDefault = true, MinWidth = 90, Margin = new(4) };
    internal ConditionalFormatRule? Result { get; private set; }
    public ConditionalRuleDialog(Window owner, CellRectangle selection, ConditionalFormatRule? existing = null)
    {
        Owner = owner; Title = "条件格式规则"; Width = 520; Height = 690; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var root = new DockPanel { Margin = new(16) }; var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(ApplyButton); buttons.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 90, Margin = new(4) });
        DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons);
        var panel = new StackPanel(); root.Children.Add(new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
        void Field(string text, Control control) { panel.Children.Add(new TextBlock { Text = text, Margin = new(0, 7, 0, 3) }); panel.Children.Add(control); }
        Field("应用于（A1:B10；多个区域用空格分隔）", Ranges); Field("规则类型", KindPicker); Field("比较方式", OperatorPicker);
        Field("条件 1 / 公式（相对引用以第一个区域首格为基准）", First); Field("条件 2（介于 / 不介于）", Second); Field("文本规则的匹配文本", MatchText);
        Field("填充色（ARGB，如 FFFFC7CE；留空不改变）", Fill); Field("字体色（ARGB；留空不改变）", Font);
        Field("数字格式代码（如 0.00%；留空不改变）", NumberFormat); panel.Children.Add(Bold); panel.Children.Add(Italic); panel.Children.Add(Stop);
        var error = new TextBlock { TextWrapping = TextWrapping.Wrap, Foreground = System.Windows.Media.Brushes.Firebrick }; panel.Children.Add(error);
        Ranges.Text = existing is null ? selection.ToString() : string.Join(" ", existing.Ranges);
        if (existing is { } v)
        {
            KindPicker.SelectedIndex = (int)v.Kind; OperatorPicker.SelectedIndex = (int)v.Comparison; First.Text = v.Formula1 ?? "";
            Second.Text = v.Formula2 ?? ""; MatchText.Text = v.Text ?? ""; Fill.Text = v.Style.FillColor ?? ""; Font.Text = v.Style.FontColor ?? "";
            Bold.IsChecked = v.Style.Bold; Italic.IsChecked = v.Style.Italic; Stop.IsChecked = v.StopIfTrue;
            NumberFormat.Text = v.Style.NumberFormatCode ?? "";
        }
        KindPicker.SelectionChanged += (_, _) => EnableFields(); OperatorPicker.SelectionChanged += (_, _) => EnableFields();
        void EnableFields()
        {
            First.IsEnabled = KindPicker.SelectedIndex is 0 or 1; OperatorPicker.IsEnabled = KindPicker.SelectedIndex == 0;
            Second.IsEnabled = KindPicker.SelectedIndex == 0 && OperatorPicker.SelectedIndex < 2; MatchText.IsEnabled = KindPicker.SelectedIndex >= 4;
        }
        EnableFields();
        ApplyButton.Click += (_, _) =>
        {
            try
            {
                var ranges = Ranges.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse).ToArray();
                if (ranges.Length == 0) throw new ArgumentException("请输入应用区域。");
                var kind = (ConditionalFormatKind)KindPicker.SelectedIndex; var first = kind is ConditionalFormatKind.CellIs or ConditionalFormatKind.Expression ? First.Text : null;
                if ((int)kind >= 4)
                {
                    var anchor = new CellRectangle(ranges[0].Top, ranges[0].Left, ranges[0].Top, ranges[0].Left).ToString(); var quoted = "\"" + MatchText.Text.Replace("\"", "\"\"") + "\"";
                    first = kind switch { ConditionalFormatKind.ContainsText => $"=NOT(ISERROR(SEARCH({quoted},{anchor})))", ConditionalFormatKind.NotContainsText => $"=ISERROR(SEARCH({quoted},{anchor}))",
                        ConditionalFormatKind.BeginsWith => $"=LEFT({anchor},LEN({quoted}))={quoted}", _ => $"=RIGHT({anchor},LEN({quoted}))={quoted}" };
                }
                string? Color(string text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim().TrimStart('#').ToUpperInvariant();
                var style = (existing?.Style ?? new()) with { Bold = Bold.IsChecked, Italic = Italic.IsChecked, FontColor = Color(Font.Text), FillColor = Color(Fill.Text),
                    NumberFormatId = NumberFormat.Text.Length == 0 ? existing?.Style.NumberFormatId : 0,
                    NumberFormatCode = NumberFormat.Text.Length == 0 ? existing?.Style.NumberFormatCode : NumberFormat.Text };
                Result = new(ranges, kind, existing?.Priority ?? 1, style, first, kind == ConditionalFormatKind.CellIs && OperatorPicker.SelectedIndex < 2 ? Second.Text : null,
                    (ConditionalComparison)OperatorPicker.SelectedIndex, Stop.IsChecked == true, (int)kind >= 4 ? MatchText.Text : null);
                DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or FormatException or NotSupportedException) { error.Text = ex.Message; }
        };
        Content = root;
    }
}

internal sealed class ConditionalFormatDialog : Window
{
    private readonly List<ConditionalFormatRule> _rules;
    internal ListBox Rules { get; } = new() { MinHeight = 190 };
    internal Button ApplyButton { get; } = new() { Content = "应用", IsDefault = true, MinWidth = 90, Margin = new(4) };
    internal IReadOnlyList<ConditionalFormatRule> Result => _rules.Select((r, i) => r.WithPriority(i + 1)).ToArray();
    public ConditionalFormatDialog(Window owner, WorksheetSnapshot sheet, CellRectangle selection)
    {
        Owner = owner; Title = "条件格式管理 · " + sheet.Name; Width = 700; Height = 420; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        _rules = sheet.Data.ConditionalFormats.OrderBy(r => r.Priority).ToList();
        var panel = new StackPanel { Margin = new(16) }; panel.Children.Add(new TextBlock { Text = "本工作表规则，按从上到下的优先级应用。修改在点击“应用”后一次提交，可撤销。", TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(Rules); var actions = new StackPanel { Orientation = Orientation.Horizontal }; panel.Children.Add(actions);
        void Button(string label, Action action) { var b = new Button { Content = label, MinWidth = 82, Margin = new(4) }; b.Click += (_, _) => action(); actions.Children.Add(b); }
        void Refresh(int selected) { Rules.ItemsSource = _rules.Select((r, i) => $"{i + 1}  {r.Kind}  {string.Join(" ", r.Ranges)}  {(r.StopIfTrue ? "为真停止" : "")}").ToArray(); Rules.SelectedIndex = selected; }
        void Edit(bool add)
        {
            var index = Rules.SelectedIndex; if (!add && index < 0) return;
            var dialog = new ConditionalRuleDialog(this, selection, add ? null : _rules[index]);
            if (dialog.ShowDialog() != true) return;
            if (add) { _rules.Insert(0, dialog.Result!); index = 0; } else _rules[index] = dialog.Result!;
            Refresh(index);
        }
        Button("新增…", () => Edit(true)); Button("编辑…", () => Edit(false));
        Button("删除", () => { var i = Rules.SelectedIndex; if (i >= 0) { _rules.RemoveAt(i); Refresh(Math.Min(i, _rules.Count - 1)); } });
        void Move(int delta) { var i = Rules.SelectedIndex; var next = i + delta; if (i < 0 || next < 0 || next >= _rules.Count) return; (_rules[i], _rules[next]) = (_rules[next], _rules[i]); Refresh(next); }
        Button("上移", () => Move(-1)); Button("下移", () => Move(1));
        panel.Children.Add(new TextBlock { Text = "支持比较、公式、重复/唯一值和文本规则。色阶、数据条和图标集仍保留在原文件中。", TextWrapping = TextWrapping.Wrap, Margin = new(0, 12, 0, 8) });
        var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        footer.Children.Add(ApplyButton); footer.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 90, Margin = new(4) });
        panel.Children.Add(footer); ApplyButton.Click += (_, _) => DialogResult = true; Content = panel; Refresh(_rules.Count == 0 ? -1 : 0);
    }
}
