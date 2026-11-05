using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task TopFilterAsync(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, WorksheetTable? table, WorksheetAutoFilter filter, int column)
    {
        var session = Document.Session.SessionId; var old = filter.Columns.GetValueOrDefault(column)?.Top;
        var dialog = new FieldsDialog(UiText.Get("前后 N 项 / 百分比"))
            .Note(UiText.Get("按完整筛选列的数值计算阈值，包含手动隐藏行，不受其他列条件影响；并列值全部保留。文本、空白和布尔不参与排名；列含错误时拒绝筛选。百分比向下取整且至少保留一项。"))
            .Choice("direction", UiText.Get("方向"), [UiText.Get("最大值优先"), UiText.Get("最小值优先")], old?.Bottom == true ? 1 : 0)
            .Text("count", UiText.Get("数量"), (old?.Count ?? 10).ToString())
            .Choice("unit", UiText.Get("单位"), [UiText.Get("项"), UiText.Get("百分比")], old?.Percent == true ? 1 : 0);
        dialog.Validate = () => dialog.Integer("count", 1, dialog.Index("unit") == 1 ? 100 : 500);
        if (!await dialog.ShowAsync(this)) return;
        if (session != Document.Session.SessionId || Document.Snapshot.ActiveSheetId != sheet.Id || Document.Snapshot.Revision != snapshot.Revision)
            throw new InvalidOperationException(UiText.Get("文档已改变，请重新预览。"));
        var rule = new AutoFilterColumn(showButton: filter.Columns.GetValueOrDefault(column)?.ShowButton ?? true,
            top: new(dialog.Integer("count", 1, dialog.Index("unit") == 1 ? 100 : 500), dialog.Index("unit") == 1, dialog.Index("direction") == 1));
        var next = new WorksheetAutoFilter(filter.Range, filter.Columns.SetItem(column, rule));
        await ExecuteAsync(table is null ? new SetAutoFilterCommand(sheet.Id, next) : new SetTableFilterCommand(sheet.Id, table.Id, next), snapshot.Revision);
        await CalculatePendingAsync(); _grid.Select(new(filter.Range.Top, column));
    }
    private async Task ColorFilterAsync(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, WorksheetTable? table, WorksheetAutoFilter filter, int column)
    {
        var session = Document.Session.SessionId; var old = filter.Columns.GetValueOrDefault(column)?.Color;
        var style = snapshot.Styles[sheet.GetStyleId(_grid.Current.Row, column)];
        var dialog = new FieldsDialog(UiText.Get("按基础颜色筛选"))
            .Note(UiText.Get("按单元格基础填充或字体颜色匹配，不使用条件格式或 Table 条纹颜色。此类筛选暂不支持含这些显示规则的列。颜色可输入 #RRGGBB 或 AARRGGBB（AA 必须为 FF）；留空表示无填充或默认黑色字体。"))
            .Choice("kind", UiText.Get("颜色类型"), [UiText.Get("填充颜色"), UiText.Get("字体颜色")], old?.Font == true ? 1 : 0)
            .Text("color", UiText.Get("筛选颜色"), old?.Argb ?? style.FillColor ?? "");
        string? Color()
        {
            var text = dialog.String("color").Trim().TrimStart('#').ToUpperInvariant();
            if (text.Length == 0) return dialog.Index("kind") == 1 ? "FF000000" : null;
            return text.Length == 6 ? "FF" + text : text;
        }
        dialog.Validate = () => _ = new AutoFilterColumn(color: new(dialog.Index("kind") == 1, Color()));
        if (!await dialog.ShowAsync(this)) return;
        if (session != Document.Session.SessionId || Document.Snapshot.ActiveSheetId != sheet.Id || Document.Snapshot.Revision != snapshot.Revision)
            throw new InvalidOperationException(UiText.Get("文档已改变，请重新预览。"));
        var rule = new AutoFilterColumn(showButton: filter.Columns.GetValueOrDefault(column)?.ShowButton ?? true, color: new(dialog.Index("kind") == 1, Color()));
        var next = new WorksheetAutoFilter(filter.Range, filter.Columns.SetItem(column, rule));
        await ExecuteAsync(table is null ? new SetAutoFilterCommand(sheet.Id, next) : new SetTableFilterCommand(sheet.Id, table.Id, next), snapshot.Revision);
        await CalculatePendingAsync(); _grid.Select(new(filter.Range.Top, column));
    }
}
