using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
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
