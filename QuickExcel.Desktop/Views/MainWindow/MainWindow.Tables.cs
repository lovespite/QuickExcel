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
            var answer = PromptText("调整表范围", $"{table.Name} 当前范围 {table.Range}。输入新的末行行号；现有单元格内容保留，含汇总行或计算列的表暂不能调整。", (table.Range.Bottom + 1).ToString(CultureInfo.InvariantCulture));
            if (answer is null) return;
            if (!int.TryParse(answer, NumberStyles.None, CultureInfo.InvariantCulture, out var bottom) || bottom < 1) throw new ArgumentException("请输入有效的末行行号。");
            if (ExecuteWorkbookCommand(new ResizeTableCommand(sheet.Id, table.Id, bottom - 1)))
            { RefreshSheet(); StatusText.Text = $"表 {table.Name} 的末行已调整为 {bottom}，单元格内容保持，可撤销。"; }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
