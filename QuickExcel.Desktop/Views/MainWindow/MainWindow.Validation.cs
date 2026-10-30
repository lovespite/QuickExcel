using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private WorksheetValidation? _smokeValidationRule;
    private bool _smokeValidationContinue;
    private bool _updatingValidation;
    private sealed record ValidationChoice(CellValue Value) { public override string ToString() => CellInput.DisplayText(Value); }
    private bool ConfirmValidation(DataValidationException exception) => _smokeMode ? _smokeValidationContinue :
        MessageBox.Show(this, exception.Message + "\n\n是否仍提交整批修改？", "数据验证", MessageBoxButton.YesNo,
            exception.Failures.All(f => f.Alert == ValidationAlert.Information) ? MessageBoxImage.Information : MessageBoxImage.Warning) == MessageBoxResult.Yes;
    private async void Validation_Click(object sender, RoutedEventArgs e) => await (_workspaceInput = EditValidationAsync());
    private async Task EditValidationAsync()
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex]; var selection = SelectedRange();
            var range = new CellRectangle(selection.Top, selection.Left, selection.Bottom, selection.Right);
            var existing = sheet.Data.Validations.Where(v => v.Ranges.Any(r => r.Top <= range.Bottom && range.Top <= r.Bottom && r.Left <= range.Right && range.Left <= r.Right)).ToArray();
            if (existing.Length > 1) throw new InvalidOperationException("选区包含多个验证规则，请选中一个规则内的单元格再修改。");
            var rule = existing.FirstOrDefault(); var ranges = rule?.Ranges.ToArray() ?? [range];
            WorksheetValidation? next;
            if (_smokeMode) next = _smokeValidationRule;
            else { var dialog = new ValidationDialog(this, ranges, rule); next = dialog.ShowDialog() == true ? dialog.Result : null; }
            if (next is null) return;
            if (ExecuteWorkbookCommand(new SetValidationCommand(sheet.Id, next)))
            { UpdateValidationHint(); StatusText.Text = "数据验证已保存，可撤销。现有内容保留；后续输入、粘贴和批量编辑使用同一规则。"; }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private async void RemoveValidation_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.Sheets[_sheetIndex]; var selected = SelectedRange(); var range = new CellRectangle(selected.Top, selected.Left, selected.Bottom, selected.Right);
            var affected = sheet.Data.Validations.Where(v => v.Ranges.Any(r => r.Top <= range.Bottom && range.Top <= r.Bottom && r.Left <= range.Right && range.Left <= r.Right)).ToArray();
            if (affected.Length == 0) { StatusText.Text = "选区没有数据验证。"; return; }
            var message = "将清除以下完整区域的验证规则，已有值保留，可撤销：\n" + string.Join("\n", affected.Select(v => string.Join(" ", v.Ranges)));
            if ((_smokeMode ? _smokeValidationContinue : MessageBox.Show(this, message, "清除数据验证", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK) &&
                ExecuteWorkbookCommand(new RemoveValidationCommand(sheet.Id, range))) UpdateValidationHint();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private void UpdateValidationHint()
    {
        if (_document is null || ValidationPanel is null) return;
        var sheet = _document.Snapshot.Sheets[_sheetIndex]; var rule = sheet.Data.Validations.FirstOrDefault(v => v.Contains(new(_selectedRow, _selectedColumn)));
        _updatingValidation = true;
        try
        {
            ValidationChoices.ItemsSource = null; ValidationChoices.SelectedIndex = -1;
            ValidationChoices.Text = "选择列表值…";
            ValidationChoices.Visibility = rule is { Kind: ValidationKind.List, ShowDropDown: true } ? Visibility.Visible : Visibility.Collapsed;
            ValidationHint.Text = rule is { ShowInput: true } ? (rule.InputTitle ?? "录入提示") + "：" + (rule.InputMessage ?? "此单元格设置了数据验证。") : "";
            ValidationHint.ToolTip = ValidationHint.Text; ValidationHint.Visibility = ValidationHint.Text.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
            ValidationPanel.Visibility = ValidationChoices.Visibility == Visibility.Visible || ValidationHint.Visibility == Visibility.Visible ? Visibility.Visible : Visibility.Collapsed;
        }
        finally { _updatingValidation = false; }
    }
    private void ValidationChoices_Opened(object sender, EventArgs e)
    {
        try
        {
            var book = _document.Snapshot; var sheet = book.Sheets[_sheetIndex]; var p = new CellPosition(_selectedRow, _selectedColumn);
            var rule = sheet.Data.Validations.First(v => v.Contains(p) && v.Kind == ValidationKind.List);
            if (_services.Calculator is not IWorkbookValidation validator) throw new NotSupportedException("数据验证计算器不可用。");
            _updatingValidation = true;
            try { ValidationChoices.ItemsSource = validator.ListValues(book, sheet.Id, p, rule).Select(v => new ValidationChoice(v)).ToArray(); }
            finally { _updatingValidation = false; }
        }
        catch (Exception ex) { ReportEditFailure(ex); ValidationChoices.IsDropDownOpen = false; }
    }
    private async void ValidationChoices_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingValidation || ValidationChoices.SelectedItem is not ValidationChoice choice || _busy) return;
        var sheetId = _document.Snapshot.Sheets[_sheetIndex].Id; var p = new CellPosition(_selectedRow, _selectedColumn);
        if (!await FlushEditsAsync()) return;
        try
        {
            var sheet = _document.Snapshot.GetSheet(sheetId);
            ExecuteWorkbookCommand(new SetCellsCommand([new(new(sheetId, p.Row, p.Column), CellContent.FromLiteral(choice.Value, sheet.GetStyleId(p.Row, p.Column)))]));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
}
