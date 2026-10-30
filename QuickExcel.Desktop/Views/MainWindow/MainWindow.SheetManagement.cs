using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private Action<NameManagerDialog>? _smokeNameManagerDriver;
    private void InstallSheetManagement()
    {
        SheetPicker.PreviewMouseRightButtonDown += (_, e) =>
        {
            if (_busy || !CanCommitComposition() || !CommitEditors()) { e.Handled = true; return; }
            if (ItemsControl.ContainerFromElement(SheetPicker, e.OriginalSource as DependencyObject) is ListBoxItem item) item.IsSelected = true;
        };
    }
    private void SheetManage_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var operation = (sender as FrameworkElement)?.Tag as string; var snapshot = _document.Snapshot; var sheet = snapshot.Sheets[_sheetIndex];
        if (operation == "copy")
        {
            var stem = sheet.Name.Length > 24 ? sheet.Name[..24] : sheet.Name; var number = 2;
            while (snapshot.Sheets.Any(s => s.Name.Equals(stem + " (" + number + ")", StringComparison.OrdinalIgnoreCase))) number++;
            var name = PromptText("复制工作表", "新工作表名称（复制到当前表之后）", stem + " (" + number + ")");
            if (name is not null) ExecuteWorkbookCommand(new CopySheetCommand(sheet.Id, name, _sheetIndex + 1));
        }
        else if (operation is "left" or "right")
        {
            var offset = operation == "left" ? -1 : 1; var target = _sheetIndex + offset;
            while (target >= 0 && target < snapshot.Sheets.Length && snapshot.Sheets[target].Visibility != SheetVisibility.Visible) target += offset;
            if (target >= 0 && target < snapshot.Sheets.Length) ExecuteWorkbookCommand(new ReorderSheetCommand(sheet.Id, target));
            else StatusText.Text = "当前方向已没有可见工作表。";
        }
    }
    private async void NameManager_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex]; var range = SelectedRange(); var revision = _document.Snapshot.Revision;
            var reference = new ReferenceNode(sheet.Name, range.Top, range.Left, range.Bottom, range.Right, true, true, true, true, range.Count > 1);
            var dialog = new NameManagerDialog(this, _document.Snapshot, "=" + FormulaSyntax.Print(reference));
            if (_smokeNameManagerDriver is { } driver) dialog.Loaded += (_, _) => Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() => driver(dialog)));
            if (dialog.ShowDialog() != true || dialog.Command is null) return;
            var result = _editor.Execute(dialog.Command, revision);
            if (!result.Changed) { StatusText.Text = "名称未改变。"; return; }
            _calculationPending = true; RefreshSheet(); UpdateChrome(); await (_gridInput = RecalculateAsync());
            StatusText.Text = "名称已更新，可一次撤销。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
