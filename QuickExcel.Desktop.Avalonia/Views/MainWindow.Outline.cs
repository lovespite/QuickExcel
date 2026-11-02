using Avalonia.Controls;
using Avalonia.Automation;
using Avalonia.Layout;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly StackPanel _outlinePanel = new() { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new(8, 0, 8, 4), IsVisible = false };
    private readonly ComboBox _outlineGroups = new() { Width = 280 };
    private sealed record OutlineItem(OutlineGroup Group)
    { public override string ToString() => $"{(Group.Axis == SheetAxis.Row ? "行" : "列")} {Group.First + 1}–{Group.Last + 1} · {Group.Level} 级 · {(Group.Collapsed ? "已折叠" : "已展开")}"; }
    private void InstallOutline()
    {
        AutomationProperties.SetName(_outlineGroups, "工作表分组"); _outlinePanel.Children.Add(_outlineGroups);
        foreach (var collapse in new[] { true, false })
        {
            var button = new Button { Content = collapse ? "− 折叠" : "+ 展开" };
            button.Click += async (_, _) => await PerformAsync(async () =>
            {
                if (_outlineGroups.SelectedItem is not OutlineItem item || !await FlushEditsAsync()) return;
                var group = item.Group; await ExecuteAsync(new SetOutlineCollapsedCommand(Document.Snapshot.ActiveSheetId, group.Axis, group.First, group.Last, group.Level, collapse));
            });
            _outlinePanel.Children.Add(button);
        }
        DockPanel.SetDock(_outlinePanel, Dock.Top); _book.Children.Insert(_book.Children.Count - 1, _outlinePanel);
    }
    private void UpdateOutline()
    {
        var selected = (_outlineGroups.SelectedItem as OutlineItem)?.Group; var sheet = Document.Snapshot.Sheets[SheetIndex];
        var items = WorksheetOutline.Groups(sheet, SheetAxis.Row).Concat(WorksheetOutline.Groups(sheet, SheetAxis.Column)).Select(g => new OutlineItem(g)).ToArray();
        _outlineGroups.ItemsSource = items;
        _outlineGroups.SelectedItem = items.FirstOrDefault(i => i.Group.Axis == selected?.Axis && i.Group.First == selected.First && i.Group.Level == selected.Level) ?? items.FirstOrDefault();
        _outlinePanel.IsVisible = items.Length > 0;
    }
    private async Task GroupAsync(bool remove)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = SelectedRange();
        var dialog = new FieldsDialog(remove ? "取消分组" : "分组").Choice("axis", "对象", [$"行 {range.Top + 1}–{range.Bottom + 1}", $"列 {range.Left + 1}–{range.Right + 1}"])
            .Note(remove ? "取消一层分组，保留数据。" : "增加一层分组，最多 7 层。");
        if (!await dialog.ShowAsync(this)) return;
        var row = dialog.Index("axis") == 0;
        await ExecuteAsync(new GroupAxisCommand(snapshot.ActiveSheetId, row ? SheetAxis.Row : SheetAxis.Column, row ? range.Top : range.Left, row ? range.Rows : range.Columns, remove), snapshot.Revision);
    }
    private async Task OutlineAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        var groups = WorksheetOutline.Groups(sheet, SheetAxis.Row).Concat(WorksheetOutline.Groups(sheet, SheetAxis.Column)).ToArray();
        if (groups.Length == 0) { _status.Text = "当前工作表没有分组。"; return; }
        var dialog = new FieldsDialog("分组展开与折叠").Choice("group", "分组", groups.Select(g => $"{(g.Axis == SheetAxis.Row ? "行" : "列")} {g.First + 1}–{g.Last + 1} · {g.Level} 级 · {(g.Collapsed ? "已折叠" : "已展开")}").ToArray())
            .Choice("state", "操作", ["展开", "折叠"]);
        if (!await dialog.ShowAsync(this)) return; var group = groups[dialog.Index("group")];
        await ExecuteAsync(new SetOutlineCollapsedCommand(sheet.Id, group.Axis, group.First, group.Last, group.Level, dialog.Index("state") == 1), snapshot.Revision);
    }
    private async Task SubtotalsAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = Rectangle(SelectedRange()); var columns = Enumerable.Range(range.Left, range.Columns).Select(QuickExcel.Formulas.FormulaSyntax.ColumnName).ToArray();
        var dialog = new FieldsDialog("分类汇总").Note("按相邻的相同分类生成汇总；请先排序。首行为标题，将先显示插入预览。")
            .Choice("group", "分类列", columns).Text("totals", "汇总列（列字母，用逗号分隔）", string.Join(',', columns.Skip(1)))
            .Choice("function", "汇总函数", ["平均值", "数值计数", "非空计数", "最大值", "最小值", "乘积", "样本标准差", "总体标准差", "求和", "样本方差", "总体方差"], 8);
        dialog.Validate = () => ColumnNumbers(dialog.String("totals")); if (!await dialog.ShowAsync(this)) return;
        var group = range.Left + dialog.Index("group"); var totals = ColumnNumbers(dialog.String("totals")); var function = dialog.Index("function") + 1;
        var preview = await Task.Run(() => SubtotalCommands.Preview(snapshot, snapshot.ActiveSheetId, range, group, totals, function));
        var confirm = new FieldsDialog("分类汇总预览").Note($"将插入 {preview.InsertedRows} 行，包含 {preview.Groups.Length} 个分类及总计。")
            .Note(string.Join("\n", preview.Groups.Select(g => $"{g.Label}：调整后第 {g.FirstRow + 1}–{g.LastRow + 1} 行 → 汇总行 {g.SummaryRow + 1}")));
        if (await confirm.ShowAsync(this)) { await ExecuteAsync(new ApplySubtotalsCommand(preview), snapshot.Revision); await CalculatePendingAsync(); }
    }
}
