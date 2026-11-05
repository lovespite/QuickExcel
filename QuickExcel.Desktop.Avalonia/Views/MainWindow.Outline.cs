using Avalonia.Controls;
using Avalonia.Automation;
using Avalonia.Layout;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void InstallOutline()
    {
        _grid.OutlineToggleRequested += group => _ = PerformAsync(async () =>
        {
            if (!await FlushEditsAsync()) return;
            if (await ExecuteAsync(new SetOutlineCollapsedCommand(Document.Snapshot.ActiveSheetId, group.Axis, group.First, group.Last, group.Level, !group.Collapsed)))
                _grid.Select(_grid.Current);
        });
        _grid.OutlineLevelRequested += level => _ = PerformAsync(async () =>
        {
            if (await FlushEditsAsync() && await ExecuteAsync(new ShowOutlineLevelCommand(Document.Snapshot.ActiveSheetId, SheetAxis.Row, level))) _grid.Select(_grid.Current);
        });
    }
    private async Task GroupAsync(bool remove)
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var range = SelectedRange();
        var dialog = new FieldsDialog(remove ? UiText.Get("取消分组") : UiText.Get("分组")).Choice("axis", UiText.Get("对象"), [UiText.Format($"行 {range.Top + 1}–{range.Bottom + 1}"), UiText.Format($"列 {range.Left + 1}–{range.Right + 1}")])
            .Note(remove ? UiText.Get("取消一层分组，保留数据。") : UiText.Get("增加一层分组，最多 7 层。"));
        if (!await dialog.ShowAsync(this)) return;
        var row = dialog.Index("axis") == 0;
        await ExecuteAsync(new GroupAxisCommand(snapshot.ActiveSheetId, row ? SheetAxis.Row : SheetAxis.Column, row ? range.Top : range.Left, row ? range.Rows : range.Columns, remove), snapshot.Revision);
    }
    private async Task OutlineAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex];
        var groups = WorksheetOutline.Groups(sheet, SheetAxis.Row).Concat(WorksheetOutline.Groups(sheet, SheetAxis.Column)).ToArray();
        if (groups.Length == 0) { _status.Text = UiText.Get("当前工作表没有分组。"); return; }
        var dialog = new FieldsDialog(UiText.Get("分组展开与折叠")).Choice("group", UiText.Get("分组"), groups.Select(g => UiText.Format($"{(g.Axis == SheetAxis.Row ? UiText.Get("行") : UiText.Get("列"))} {g.First + 1}–{g.Last + 1} · {g.Level} 级 · {(g.Collapsed ? UiText.Get("已折叠") : UiText.Get("已展开"))}")).ToArray())
            .Choice("state", UiText.Get("操作"), [UiText.Get("展开"), UiText.Get("折叠")]);
        if (!await dialog.ShowAsync(this)) return; var group = groups[dialog.Index("group")];
        await ExecuteAsync(new SetOutlineCollapsedCommand(sheet.Id, group.Axis, group.First, group.Last, group.Level, dialog.Index("state") == 1), snapshot.Revision);
    }
    private async Task SubtotalsAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var range = DataToolRange();
        if (range.Columns < 2 || range.Rows < 2) throw new ArgumentException(UiText.Get("分类汇总需要包含标题、分类列和汇总列的数据区域。"));
        var columns = Enumerable.Range(range.Left, range.Columns).Select(c => DataColumnLabel(sheet, range.Top, c)).ToArray();
        var selected = Enumerable.Range(0, columns.Length).Select(i => new CheckBox { Content = columns[i], IsChecked = i > 0 }).ToArray();
        var checklist = new StackPanel { Spacing = 6 };
        foreach (var check in selected) checklist.Children.Add(check);
        var remove = false;
        var dialog = new FieldsDialog(UiText.Get("分类汇总")).Note(UiText.Get("按相邻的相同分类生成汇总；请先排序。首行为标题，将先显示插入预览。"))
            .Choice("group", UiText.Get("分类字段"), columns)
            .Choice("function", UiText.Get("汇总方式"), [UiText.Get("平均值"), UiText.Get("数值计数"), UiText.Get("非空计数"), UiText.Get("最大值"), UiText.Get("最小值"), UiText.Get("乘积"), UiText.Get("样本标准差"), UiText.Get("总体标准差"), UiText.Get("求和"), UiText.Get("样本方差"), UiText.Get("总体方差")], 8)
            .Note(UiText.Get("选定汇总项：")).Preview(new ScrollViewer { Content = checklist, MinHeight = 110, MaxHeight = 220 })
            .Check("replace", UiText.Get("替换当前分类汇总"), true).Check("pages", UiText.Get("每组数据分页")).Check("below", UiText.Get("汇总结果显示在数据下方"), sheet.Data.Outline?.SummaryBelow ?? true)
            .Alternative(UiText.Get("全部删除"), () => remove = true);
        dialog.Validate = () =>
        {
            if (!selected.Any(c => c.IsChecked == true)) throw new ArgumentException(UiText.Get("请至少选择一个汇总项。"));
            if (selected[dialog.Index("group")].IsChecked == true) throw new ArgumentException(UiText.Get("汇总项不能与分类字段相同。"));
        };
        if (!await dialog.ShowAsync(this)) return;
        var group = range.Left + dialog.Index("group"); var totals = selected.Select((c, i) => (c, i)).Where(p => p.c.IsChecked == true).Select(p => range.Left + p.i).ToArray(); var function = dialog.Index("function") + 1;
        var options = new SubtotalOptions(dialog.Boolean("replace"), dialog.Boolean("pages"), dialog.Boolean("below"));
        var preview = await Task.Run(() => remove ? SubtotalCommands.Remove(snapshot, sheet.Id, range) : SubtotalCommands.Preview(snapshot, sheet.Id, range, group, totals, function, options));
        var confirm = new FieldsDialog(UiText.Get("分类汇总预览")).Note(remove ? UiText.Format($"将删除 {preview.RemovedRows} 行分类汇总并保留明细数据，可整体撤销。") :
            UiText.Format($"将删除旧汇总 {preview.RemovedRows} 行、插入 {preview.InsertedRows} 行，包含 {preview.Groups.Length} 个分类及总计。"))
            .Note(string.Join("\n", preview.Groups.Take(30).Select(g => UiText.Format($"{g.Label}：第 {g.FirstRow + 1}–{g.LastRow + 1} 行 → 汇总行 {g.SummaryRow + 1}"))));
        if (await confirm.ShowAsync(this) && await ExecuteAsync(new ApplySubtotalsCommand(preview), snapshot.Revision))
        {
            await CalculatePendingAsync(); var result = preview.Range;
            _grid.SelectRange(new(result.Top, result.Left, result.Bottom, result.Right), new(result.Top, result.Left));
        }
    }
}
