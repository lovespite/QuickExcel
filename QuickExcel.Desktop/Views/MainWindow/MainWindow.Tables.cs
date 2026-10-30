using System.Globalization;
using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async void CreateTable_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = CreateTableAsync());
    private async Task CreateTableAsync()
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var range = SelectedRange(); var snapshot = _document.Snapshot; var sheet = snapshot.Sheets[_sheetIndex];
            var suffix = 1; while (snapshot.Sheets.SelectMany(s => s.Data.Tables).Any(t => t.Name.Equals("QE_Table" + suffix, StringComparison.OrdinalIgnoreCase))) suffix++;
            var name = PromptText("创建表", "所选区域首行须为非空、唯一的文本表头。请输入工作簿内唯一的表名称：", "QE_Table" + suffix);
            if (name is null) return;
            var style = TableStyles.Resolve(snapshot, "TableStyleMedium2");
            if (ExecuteWorkbookCommand(new CreateTableCommand(sheet.Id, new(range.Top, range.Left, range.Bottom, range.Right), name, style)))
            { RefreshSheet(); StatusText.Text = $"已创建表 {name}；表头、筛选和隔行底色随文件保存，可撤销。"; }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private async void ResizeTable_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = ResizeTableAsync());
    private async Task ResizeTableAsync()
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex];
            var table = sheet.Data.Tables.FirstOrDefault(t => t.Range.Contains(_selectedRow, _selectedColumn)) ?? throw new InvalidOperationException("请先选择表内单元格。");
            var answer = PromptText("调整表范围", $"{table.Name} 当前范围 {table.Range}。输入新的末行行号；现有内容保留，新行续填计算列。含汇总行时请在表内插入或删除行。", (table.Range.Bottom + 1).ToString(CultureInfo.InvariantCulture));
            if (answer is null) return;
            if (!int.TryParse(answer, NumberStyles.None, CultureInfo.InvariantCulture, out var bottom) || bottom < 1) throw new ArgumentException("请输入有效的末行行号。");
            if (ExecuteWorkbookCommand(new ResizeTableCommand(sheet.Id, table.Id, bottom - 1)))
            { RefreshSheet(); StatusText.Text = $"表 {table.Name} 的末行已调整为 {bottom}，现有内容保留，新行计算列已续填，可撤销。"; }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private async void RenameTable_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = EditTableDefinitionAsync(0));
    private async void RenameTableColumn_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = EditTableDefinitionAsync(1));
    private async void CalculatedTableColumn_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = EditTableDefinitionAsync(2));
    private async Task EditTableDefinitionAsync(int action)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex];
            var table = sheet.Data.Tables.FirstOrDefault(t => t.Range.Contains(_selectedRow, _selectedColumn)) ?? throw new InvalidOperationException("请先选择表内单元格。");
            var column = table.Columns[_selectedColumn - table.Range.Left];
            var answer = action switch
            {
                0 => PromptText("重命名表", "所有绑定到此表的公式会同步更新。请输入新的表名：", table.Name),
                1 => PromptText("重命名表列", "表头和引用此列的公式会同步更新。请输入新的列名：", column.Name),
                _ => PromptText("设置计算列", $"将替换“{column.Name}”整列数据区，包括已有例外值；整列替换可一次撤销。输入以 = 开头的公式，例如 =[@数量]*[@单价]：",
                    column.CalculatedFormula ?? "=")
            };
            if (answer is null) return;
            DocumentCommand command = action switch
            {
                0 => new RenameTableCommand(sheet.Id, table.Id, answer),
                1 => new RenameTableColumnCommand(sheet.Id, table.Id, column.Id, answer),
                _ => new SetTableCalculatedColumnCommand(sheet.Id, table.Id, column.Id, answer)
            };
            if (ExecuteWorkbookCommand(command))
            { RefreshSheet(); StatusText.Text = action == 2 ? "计算列已更新，可一次撤销。单格编辑只改变当前格，不自动覆盖其他行。" : "表定义与引用公式已同步更新，可撤销。"; }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
