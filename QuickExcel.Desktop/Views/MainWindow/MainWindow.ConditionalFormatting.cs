using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private IReadOnlyList<ConditionalFormatRule>? _smokeConditionalFormats;
    private async void ConditionalFormatting_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = EditConditionalFormatsAsync());
    private async Task EditConditionalFormatsAsync()
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex]; var selected = SelectedRange();
            IReadOnlyList<ConditionalFormatRule>? rules;
            if (_smokeMode) rules = _smokeConditionalFormats;
            else
            {
                var dialog = new ConditionalFormatDialog(this, sheet, new(selected.Top, selected.Left, selected.Bottom, selected.Right));
                rules = dialog.ShowDialog() == true ? dialog.Result : null;
            }
            if (rules is not null && ExecuteWorkbookCommand(new SetConditionalFormatsCommand(sheet.Id, rules)))
                StatusText.Text = "条件格式已更新，可撤销。规则按优先级叠加，编辑数据后自动刷新。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
