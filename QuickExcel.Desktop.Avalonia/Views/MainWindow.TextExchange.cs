using QuickExcel.Core;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task ExportTextAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot;
        var sheetId = snapshot.ActiveSheetId;
        var session = Document.Session.SessionId;
        var selection = Rectangle(SelectedRange());
        var dialog = new FieldsDialog(UiText.Get("导出 CSV / TSV"))
            .Note(UiText.Get("仅导出当前表或矩形选区的值，包含隐藏行列；不保留格式、其他工作表或对象。公式必须具有最新结果。原始值中的日期为序列数；显示文本使用当前数字和日期格式。"))
            .Choice("scope", UiText.Get("导出范围"), [UiText.Get("当前工作表"), UiText.Get("当前选区")])
            .Choice("separator", UiText.Get("分隔符"), [UiText.Get("逗号（CSV）"), UiText.Get("制表符（TSV）")])
            .Choice("values", UiText.Get("导出内容"), [UiText.Get("原始值"), UiText.Get("显示文本")])
            .Check("bom", UiText.Get("UTF-8 BOM（便于 Excel 识别中文）"), true);
        if (!await dialog.ShowAsync(this)) return;
        var range = dialog.Index("scope") == 0 ? WorkbookTextExchange.UsedRange(snapshot.GetSheet(sheetId)) : selection;
        var options = new TextExportOptions(dialog.Index("separator") == 0 ? ',' : '\t', dialog.Boolean("bom"), dialog.Index("values") == 1);
        var extension = options.Separator == ',' ? ".csv" : ".tsv";
        var file = await _files.SaveAsync(new(UiText.Get("导出 CSV / TSV"), [new("CSV / TSV", ["*" + extension])], snapshot.GetSheet(sheetId).Name + extension, extension));
        if (file is null) return;
        if (session != Document.Session.SessionId || Document.Snapshot.Revision != snapshot.Revision || Document.Snapshot.ActiveSheetId != sheetId)
            throw new InvalidOperationException(UiText.Get("文档已改变，请重新预览。"));
        if (_workspace.FilePath is { } source && string.Equals(Path.GetFullPath(file.Path), Path.GetFullPath(source), OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new InvalidOperationException(UiText.Get("导出需要使用新副本路径，不能覆盖当前打开的工作簿。"));
        using var cancellation = new CancellationTokenSource();
        using var feedback = BeginExecution(UiText.Get("正在导出文本…"), cancellation);
        await Task.Run(() =>
        {
            var version = FileVersions.Read(file.Path);
            var bytes = WorkbookTextExchange.Export(snapshot, sheetId, range, options, cancellation.Token);
            WorkbookTextExchange.Save(file.Path, bytes, version, cancellation.Token);
        }, cancellation.Token);
        _status.Text = UiText.Get("已导出 ") + file.Path;
    }
}
