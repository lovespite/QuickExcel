using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private static readonly string[] ConditionalKinds = ["单元格比较", "公式", "重复值", "唯一值", "包含文本", "不包含文本", "以文本开头", "以文本结尾", "色阶", "数据条", "图标集"];
    private async Task ConditionalFormatsAsync()
    {
        if (!await FlushEditsAsync()) return;
        RequireConditionalFormatEditing();
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var selection = Rectangle(SelectedRange());
        var rules = sheet.Data.ConditionalFormats.OrderBy(r => r.Priority).ToList();
        var window = new Window { Title = "条件格式管理 · " + sheet.Name, Width = 720, Height = 440, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        var panel = new StackPanel { Margin = new(20), Spacing = 12 };
        panel.Children.Add(new TextBlock { Text = "从上到下按优先级应用。点击“应用”后一次提交，可撤销。", TextWrapping = TextWrapping.Wrap });
        var list = new ListBox { Height = 230 }; panel.Children.Add(list);
        var buttons = new WrapPanel(); panel.Children.Add(buttons);
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap }; panel.Children.Add(error);
        void Refresh(int index) { list.ItemsSource = rules.Select((r, i) => $"{i + 1} · {ConditionalKinds[(int)r.Kind]} · {string.Join(" ", r.Ranges)}{(r.StopIfTrue ? " · 为真停止" : "")}").ToArray(); list.SelectedIndex = index; }
        void Button(string name, Func<Task> action)
        {
            var button = new Button { Content = name, Margin = new(3) }; buttons.Children.Add(button);
            button.Click += async (_, _) => { buttons.IsEnabled = false; try { await action(); } catch (Exception ex) { error.Text = ex.Message; } finally { buttons.IsEnabled = true; } };
        }
        async Task Edit(bool add)
        {
            var index = list.SelectedIndex; if (!add && index < 0) return;
            var old = add ? null : rules[index]; var dialog = RuleFields(selection, old, out var read);
            dialog.Validate = () => read().ValidateBinding(snapshot, sheet.Id);
            if (!await dialog.ShowAsync(window)) return;
            if (add) { rules.Insert(0, read()); index = 0; } else rules[index] = read(); Refresh(index);
        }
        Button("新增", () => Edit(true)); Button("编辑", () => Edit(false));
        Button("删除", () => { var i = list.SelectedIndex; if (i >= 0) { rules.RemoveAt(i); Refresh(Math.Min(i, rules.Count - 1)); } return Task.CompletedTask; });
        Task Move(int delta) { var i = list.SelectedIndex; if (i >= 0 && i + delta >= 0 && i + delta < rules.Count) { (rules[i], rules[i + delta]) = (rules[i + delta], rules[i]); Refresh(i + delta); } return Task.CompletedTask; }
        Button("上移", () => Move(-1)); Button("下移", () => Move(1));
        Button("应用", () => { window.Close(true); return Task.CompletedTask; }); Button("取消", () => { window.Close(false); return Task.CompletedTask; });
        Refresh(rules.Count > 0 ? 0 : -1); window.Content = panel;
        if (await ShowOwnedAsync<bool>(window)) await ExecuteAsync(new SetConditionalFormatsCommand(sheet.Id, rules.Select((r, i) => r.WithPriority(i + 1))), snapshot.Revision);
    }
    private static FieldsDialog RuleFields(CellRectangle selection, ConditionalFormatRule? old, out Func<ConditionalFormatRule> read)
    {
        if (old?.Kind >= ConditionalFormatKind.ColorScale) return VisualRuleFields(old, out read);
        static int Tri(bool? value) => value is null ? 0 : value.Value ? 1 : 2;
        var dialog = new FieldsDialog("条件格式规则").Text("ranges", "应用区域（多个区域用空格分隔）", old is null ? selection.ToString() : string.Join(" ", old.Ranges))
            .Choice("kind", "规则类型", ConditionalKinds.Take(8).ToArray(), (int)(old?.Kind ?? ConditionalFormatKind.CellIs))
            .Choice("comparison", "比较方式", ["介于", "不介于", "等于", "不等于", "小于", "小于等于", "大于", "大于等于"], (int)(old?.Comparison ?? ConditionalComparison.Between))
            .Text("first", "条件 1 / 公式（相对引用以首格为基准）", old?.Formula1 ?? "=1").Text("second", "条件 2", old?.Formula2 ?? "=10")
            .Text("text", "文本条件", old?.Text ?? "").Text("fill", "填充色（ARGB，留空不改变）", old is null ? "FFFFC7CE" : old.Style.FillColor ?? "")
            .Text("font", "字体色（ARGB，留空不改变）", old is null ? "FF9C0006" : old.Style.FontColor ?? "").Text("number", "数字格式代码（留空保持）", old?.Style.NumberFormatCode ?? "")
            .Choice("bold", "粗体", ["不改变", "是", "否"], Tri(old?.Style.Bold)).Choice("italic", "斜体", ["不改变", "是", "否"], Tri(old?.Style.Italic))
            .Check("stop", "为真时停止后续规则", old?.StopIfTrue ?? false);
        read = () =>
        {
            var ranges = dialog.String("ranges").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse).ToArray();
            if (ranges.Length == 0) throw new ArgumentException("请输入应用区域。");
            var kind = (ConditionalFormatKind)dialog.Index("kind"); var comparison = (ConditionalComparison)dialog.Index("comparison");
            var first = kind is ConditionalFormatKind.CellIs or ConditionalFormatKind.Expression ? dialog.String("first") : null;
            if ((int)kind >= 4)
            {
                var anchor = new CellRectangle(ranges[0].Top, ranges[0].Left, ranges[0].Top, ranges[0].Left).ToString();
                var quoted = "\"" + dialog.String("text").Replace("\"", "\"\"") + "\"";
                first = kind switch { ConditionalFormatKind.ContainsText => $"=NOT(ISERROR(SEARCH({quoted},{anchor})))", ConditionalFormatKind.NotContainsText => $"=ISERROR(SEARCH({quoted},{anchor}))",
                    ConditionalFormatKind.BeginsWith => $"=LEFT({anchor},LEN({quoted}))={quoted}", _ => $"=RIGHT({anchor},LEN({quoted}))={quoted}" };
            }
            string? Color(string key) => string.IsNullOrWhiteSpace(dialog.String(key)) ? null : dialog.String(key).Trim().TrimStart('#').ToUpperInvariant();
            bool? Flag(string key) => dialog.Index(key) == 0 ? null : dialog.Index(key) == 1;
            var style = (old?.Style ?? new()) with { Bold = Flag("bold"), Italic = Flag("italic"), FontColor = Color("font"), FillColor = Color("fill"),
                NumberFormatId = dialog.String("number").Length == 0 ? old?.Style.NumberFormatId : 0,
                NumberFormatCode = dialog.String("number").Length == 0 ? old?.Style.NumberFormatCode : dialog.String("number") };
            return new(ranges, kind, old?.Priority ?? 1, style, first, kind == ConditionalFormatKind.CellIs && comparison is ConditionalComparison.Between or ConditionalComparison.NotBetween ? dialog.String("second") : null,
                comparison, dialog.Boolean("stop"), (int)kind >= 4 ? dialog.String("text") : null) { ScriptId = old?.ScriptId ?? Guid.NewGuid() };
        };
        return dialog;
    }
}
