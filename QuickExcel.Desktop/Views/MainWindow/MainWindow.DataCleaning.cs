using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private CleaningPreview? _smokeCleaningPreview;
    private bool _smokeCleaningOverwrite;
    private async void RemoveDuplicates_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = CleanDataAsync(CleaningOperation.Duplicates));
    private async void SplitText_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = CleanDataAsync(CleaningOperation.Split));
    private async void ConvertNumbers_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = CleanDataAsync(CleaningOperation.Numbers));
    private async Task CleanDataAsync(CleaningOperation operation)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var book = _document.Snapshot; var sheet = book.Sheets[_sheetIndex]; var selected = SelectedRange();
            CurrentView.RequireCurrentFilter(book);
            CleaningPreview? preview; bool overwrite;
            if (_smokeMode) { preview = _smokeCleaningPreview; overwrite = _smokeCleaningOverwrite; }
            else
            {
                var dialog = new CleaningDialog(this, book, sheet.Id, new(selected.Top, selected.Left, selected.Bottom, selected.Right), operation);
                if (dialog.ShowDialog() != true) return; preview = dialog.Preview; overwrite = dialog.AllowOverwrite;
            }
            if (preview is not null && ExecuteWorkbookCommand(new ApplyCleaningCommand(preview, overwrite)))
                StatusText.Text = $"清洗已应用：修改 {preview.Changes.Length} 格，删除重复记录 {preview.RemovedRows.Length} 行；可一次撤销。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
