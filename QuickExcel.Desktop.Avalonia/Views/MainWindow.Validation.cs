using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly StackPanel _validationPanel = new() { Orientation = Orientation.Horizontal, Spacing = 10, Margin = new(8, 0, 8, 4), IsVisible = false };
    private readonly ComboBox _validationValues = new() { MinWidth = 170, PlaceholderText = UiText.Get("选择列表值") };
    private readonly TextBlock _validationHint = new() { VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
    private bool _updatingValidation;
    private sealed record ValidationValue(CellValue Value) { public override string ToString() => CellInput.DisplayText(Value); }
    private void InstallValidation()
    {
        _validationPanel.Children.Add(_validationValues); _validationPanel.Children.Add(_validationHint);
        _validationValues.DropDownOpened += (_, _) =>
        {
            try
            {
                var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var rule = sheet.Data.Validations.First(v => v.Contains(_grid.Current) && v.Kind == ValidationKind.List);
                if (_services.Calculator is not IWorkbookValidation validation) throw new NotSupportedException(UiText.Get("数据验证计算器不可用。"));
                _updatingValidation = true; _validationValues.ItemsSource = validation.ListValues(snapshot, sheet.Id, _grid.Current, rule).Select(v => new ValidationValue(v)).ToArray();
            }
            catch (Exception ex) { Error(ex); _validationValues.IsDropDownOpen = false; }
            finally { _updatingValidation = false; }
        };
        _validationValues.SelectionChanged += async (_, _) =>
        {
            if (_updatingValidation || _validationValues.SelectedItem is not ValidationValue value) return;
            var sheet = Document.Snapshot.ActiveSheetId; var owner = _grid.Current;
            await PerformAsync(async () =>
            {
                if (!await CommitEditorsAsync()) return;
                await ExecuteAsync(new SetCellsCommand([new(new(sheet, owner.Row, owner.Column), CellContent.FromLiteral(value.Value, Document.Snapshot.GetSheet(sheet).GetStyleId(owner.Row, owner.Column)))]));
            });
        };
    }
    private void UpdateValidation()
    {
        var rule = Document.Snapshot.Sheets[SheetIndex].Data.Validations.FirstOrDefault(v => v.Contains(_grid.Current));
        _updatingValidation = true;
        try
        {
            _validationValues.ItemsSource = null; _validationValues.SelectedIndex = -1;
            _validationValues.IsVisible = rule is { Kind: ValidationKind.List, ShowDropDown: true };
            _validationHint.Text = rule is { ShowInput: true } ? (rule.InputTitle ?? UiText.Get("录入提示")) + "：" + (rule.InputMessage ?? UiText.Get("此单元格设置了数据验证。")) : "";
            ToolTip.SetTip(_validationHint, _validationHint.Text); _validationHint.IsVisible = !string.IsNullOrEmpty(_validationHint.Text);
            _validationPanel.IsVisible = _validationValues.IsVisible || _validationHint.IsVisible;
        }
        finally { _updatingValidation = false; }
    }
    private async Task EditValidationAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var selected = Rectangle(SelectedRange());
        var matches = sheet.Data.Validations.Where(v => v.Ranges.Any(r => r.Top <= selected.Bottom && r.Bottom >= selected.Top && r.Left <= selected.Right && r.Right >= selected.Left)).ToArray();
        if (matches.Length > 1) throw new InvalidOperationException(UiText.Get("选区包含多个验证规则，请选择一个规则内的单元格。"));
        var old = matches.FirstOrDefault(); var ranges = old?.Ranges.ToArray() ?? [selected];
        var dialog = new FieldsDialog(UiText.Get("数据验证")).Note(UiText.Get("应用范围：") + string.Join(" ", ranges.Select(r => r.ToString())))
            .Choice("kind", UiText.Get("允许"), [UiText.Get("任何值"), UiText.Get("整数"), UiText.Get("小数"), UiText.Get("列表"), UiText.Get("日期"), UiText.Get("文本长度"), UiText.Get("自定义公式")], (int)(old?.Kind ?? ValidationKind.Whole))
            .Choice("operator", UiText.Get("条件"), [UiText.Get("介于"), UiText.Get("不介于"), UiText.Get("等于"), UiText.Get("不等于"), UiText.Get("小于"), UiText.Get("小于等于"), UiText.Get("大于"), UiText.Get("大于等于")], (int)(old?.Comparison ?? ValidationOperator.Between))
            .Text("first", UiText.Get("条件 1 / 列表来源 / 公式"), old?.Formula1 ?? "=1").Text("second", UiText.Get("条件 2（介于或不介于）"), old?.Formula2 ?? "=10")
            .Check("blank", UiText.Get("允许空白"), old?.AllowBlank ?? true).Check("dropdown", UiText.Get("显示列表下拉框"), old?.ShowDropDown ?? true)
            .Check("input", UiText.Get("显示输入提示"), old?.ShowInput ?? true).Text("inputTitle", UiText.Get("提示标题"), old?.InputTitle ?? "").Text("inputText", UiText.Get("提示内容"), old?.InputMessage ?? "")
            .Check("error", UiText.Get("显示错误提示"), old?.ShowError ?? true).Choice("alert", UiText.Get("错误样式"), [UiText.Get("停止（拒绝输入）"), UiText.Get("警告（允许确认）"), UiText.Get("信息（允许确认）")], (int)(old?.Alert ?? ValidationAlert.Stop))
            .Text("errorTitle", UiText.Get("错误标题"), old?.ErrorTitle ?? "").Text("errorText", UiText.Get("错误内容"), old?.ErrorMessage ?? "");
        WorksheetValidation Read()
        {
            var kind = (ValidationKind)dialog.Index("kind"); var comparison = (ValidationOperator)dialog.Index("operator");
            string? Optional(string key) => string.IsNullOrEmpty(dialog.String(key)) ? null : dialog.String(key);
            return new(ranges, kind, kind == ValidationKind.None ? null : dialog.String("first"),
                kind is ValidationKind.Whole or ValidationKind.Decimal or ValidationKind.Date or ValidationKind.TextLength && comparison is ValidationOperator.Between or ValidationOperator.NotBetween ? dialog.String("second") : null,
                comparison, dialog.Boolean("blank"), dialog.Boolean("dropdown"), dialog.Boolean("input"), dialog.Boolean("error"), (ValidationAlert)dialog.Index("alert"),
                Optional("inputTitle"), Optional("inputText"), Optional("errorTitle"), Optional("errorText"));
        }
        dialog.Validate = () => Read().ValidateBinding(snapshot, sheet.Id);
        if (await dialog.ShowAsync(this)) await ExecuteAsync(new SetValidationCommand(sheet.Id, Read()), snapshot.Revision);
    }
    private async Task RemoveValidationAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[SheetIndex]; var range = Rectangle(SelectedRange());
        var affected = sheet.Data.Validations.Where(v => v.Ranges.Any(r => r.Top <= range.Bottom && r.Bottom >= range.Top && r.Left <= range.Right && r.Right >= range.Left)).ToArray();
        if (affected.Length == 0) return;
        if (await Dialogs.Confirm(this, UiText.Get("清除数据验证"), UiText.Get("将清除以下完整区域的规则，已有值保留，可撤销：\n") + string.Join("\n", affected.Select(v => string.Join(" ", v.Ranges)))))
            await ExecuteAsync(new RemoveValidationCommand(sheet.Id, range), snapshot.Revision);
    }
}
