using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async void ChartManage_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var selected = SelectedRange();
        var dialog = new ReportFieldsDialog(this, "图表：插入、数据范围和标题");
        var list = dialog.Choice("图表", new[] { "插入新图表" }.Concat(sheet.Data.Charts.Select((c, i) => $"{i + 1} · {c.Title}")).ToArray(), 0);
        var type = dialog.Choice("类型", ["簇状柱形图", "折线图", "饼图"], 0);
        var title = dialog.Text("图表标题", "");
        var source = dialog.Text("数据范围：首行为系列标题，首列为分类；编辑时留空保持原范围", new CellRectangle(selected.Top, selected.Left, selected.Bottom, selected.Right).ToString());
        var legend = dialog.Choice("图例", ["右侧", "底部", "左侧", "顶部", "隐藏"], 0);
        var remove = dialog.Check("删除所选图表", false);
        list.SelectionChanged += (_, _) =>
        {
            if (list.SelectedIndex <= 0) return;
            var chart = sheet.Data.Charts[list.SelectedIndex - 1]; title.Text = chart.Title; type.SelectedIndex = (int)chart.Kind;
            legend.SelectedIndex = chart.Legend switch { "r" => 0, "b" => 1, "l" => 2, "t" => 3, _ => 4 }; source.Text = "";
        };
        dialog.Apply = () =>
        {
            var index = list.SelectedIndex <= 0 ? sheet.Data.Charts.Length : list.SelectedIndex - 1;
            if (remove.IsChecked == true) { ExecuteWorkbookCommand(new SetChartCommand(sheet.Id, index, null)); return; }
            var old = index < sheet.Data.Charts.Length ? sheet.Data.Charts[index] : null;
            var anchor = old?.Anchor ?? DrawingGeometry.FromBounds(sheet, new(DrawingGeometry.ColumnStart(sheet, selected.Left), DrawingGeometry.RowStart(sheet, selected.Bottom + 2), 560, 320), DrawingAnchorKind.OneCell);
            var chart = string.IsNullOrWhiteSpace(source.Text) && old is not null ? old with { Kind = (WorksheetChartKind)type.SelectedIndex, Title = title.Text } : SetChartCommand.FromRange(sheet, CellRectangle.Parse(source.Text), (WorksheetChartKind)type.SelectedIndex, anchor, title.Text);
            chart = chart with { Legend = legend.SelectedIndex switch { 0 => "r", 1 => "b", 2 => "l", 3 => "t", _ => null } };
            ExecuteWorkbookCommand(new SetChartCommand(sheet.Id, index, chart));
        };
        dialog.ShowDialog();
    }

    private async void Protection_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var old = sheet.Data.Protection;
        var dialog = new ReportFieldsDialog(this, "工作表保护（操作约束，不是文件加密）");
        var enabled = dialog.Check("启用工作表保护（不设置密码）", old?.Sheet == true);
        var cells = dialog.Check("允许设置单元格格式", old?.FormatCells == false); var rows = dialog.Check("允许设置行格式和尺寸", old?.FormatRows == false); var columns = dialog.Check("允许设置列格式和尺寸", old?.FormatColumns == false);
        var insertRows = dialog.Check("允许插入行", old?.InsertRows == false); var insertColumns = dialog.Check("允许插入列", old?.InsertColumns == false);
        var filter = dialog.Check("允许调整现有筛选条件", old?.AutoFilter == false); var sort = dialog.Check("允许排序未锁定区域", old?.Sort == false); var objects = dialog.Check("允许编辑图片、图表和传统批注", old?.Objects == false);
        var selectLocked = dialog.Check("允许选择锁定单元格", old?.SelectLockedCells != true); var selectUnlocked = dialog.Check("允许选择未锁定单元格", old?.SelectUnlockedCells != true);
        dialog.Apply = () => ExecuteWorkbookCommand(new SetProtectionCommand(sheet.Id, enabled.IsChecked == true ? (old ?? new WorksheetProtection()) with
        {
            Sheet = true,
            FormatCells = cells.IsChecked != true, FormatRows = rows.IsChecked != true, FormatColumns = columns.IsChecked != true, InsertRows = insertRows.IsChecked != true,
            InsertColumns = insertColumns.IsChecked != true, AutoFilter = filter.IsChecked != true, Sort = sort.IsChecked != true, Objects = objects.IsChecked != true,
            SelectLockedCells = selectLocked.IsChecked != true, SelectUnlockedCells = selectUnlocked.IsChecked != true
        } : null));
        dialog.ShowDialog();
    }
    private async void CellProtection_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var range = SelectedRange(); var style = _document.Snapshot.Styles[sheet.GetStyleId(range.Top, range.Left)];
        var dialog = new ReportFieldsDialog(this, "选区锁定属性");
        var locked = dialog.Check("锁定：工作表保护后禁止修改内容", style.Locked); var hidden = dialog.Check("工作表保护后隐藏公式", style.FormulaHidden);
        dialog.Apply = () => ExecuteWorkbookCommand(new SetCellProtectionCommand(sheet.Id, new(range.Top, range.Left, range.Bottom, range.Right), locked.IsChecked == true, hidden.IsChecked == true)); dialog.ShowDialog();
    }
    private async void NoteEdit_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var position = new CellPosition(_selectedRow, _selectedColumn); var note = sheet.Annotations.Notes.GetValueOrDefault(position);
        var dialog = new ReportFieldsDialog(this, "传统批注"); var author = dialog.Text("作者", note?.Author ?? "QuickExcel");
        var text = dialog.Field("批注内容", new TextBox { Text = note?.Text.Text ?? "", AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, MinHeight = 220, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
        var flatten = dialog.Check("替换含格式的批注时，明确转换为普通文本", false); var remove = dialog.Check("删除此批注", false);
        dialog.Apply = () => ExecuteWorkbookCommand(new SetNoteCommand(sheet.Id, position, remove.IsChecked == true ? null : new CellNote(author.Text, new([new TextRun(text.Text)])), flatten.IsChecked == true)); dialog.ShowDialog();
    }
}
