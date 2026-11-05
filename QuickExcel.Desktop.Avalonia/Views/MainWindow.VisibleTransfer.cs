using Avalonia;
using Avalonia.Controls;
using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed record VisibleCapture(string Token, Guid Session, long Revision, ExternalClipboardRegion Data);
    private VisibleCapture? _visibleCapture;
    private void InstallVisibleTransfer()
    {
        var edit = (MenuItem)_menu.Items[1]!;
        foreach (var (label, action) in new (string, Func<Task>)[] { (UiText.Get("复制可见单元格…"), CopyVisibleAsync), (UiText.Get("粘贴到可见单元格…"), PasteVisibleAsync) })
        {
            var item = new MenuItem { Header = label }; item.Bind(IsEnabledProperty, _book.GetObservable(IsVisibleProperty));
            item.Click += async (_, _) => { if (_book.IsVisible) await PerformAsync(action); }; edit.Items.Add(item);
        }
    }
    private async Task CopyVisibleAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var target = SelectedRange();
        var dialog = new FieldsDialog(UiText.Get("复制可见单元格"))
            .Note(UiText.Get("跳过隐藏和筛除行，保留列位置、类型和基本格式；不压缩隐藏列。富文本、链接、批注、合并及未映射规则须先移出选区。公式模式仅支持普通相对引用，拒绝数组、名称、跨表及绝对引用。"))
            .Note(UiText.Get("公式模式仅发布 XML 表格格式，需要目标应用支持该格式。"))
            .Choice("mode", UiText.Get("复制内容"), [UiText.Get("结果值和基本格式"), UiText.Get("公式（相对引用）")]);
        if (!await dialog.ShowAsync(this)) return;
        if (session != Document.Session.SessionId || snapshot.Revision != Document.Snapshot.Revision || target != SelectedRange() || snapshot.ActiveSheetId != Document.Snapshot.ActiveSheetId)
            throw new InvalidOperationException(UiText.Get("文档或选区已改变，请重新定位。"));
        var formulas = dialog.Index("mode") == 1;
        using var cancellation = new CancellationTokenSource(); ExternalClipboardRegion region;
        using (BeginExecution(UiText.Get("正在复制可见单元格…"), cancellation))
            region = await Task.Run(() => VisibleCellTransfer.Capture(snapshot, snapshot.ActiveSheetId, Rectangle(target), formulas, cancellation.Token), cancellation.Token);
        if (session != Document.Session.SessionId || snapshot.Revision != Document.Snapshot.Revision || target != SelectedRange() || snapshot.ActiveSheetId != Document.Snapshot.ActiveSheetId ||
            !ReferenceEquals(snapshot.GetSheet(snapshot.ActiveSheetId), Document.Snapshot.GetSheet(snapshot.ActiveSheetId)))
            throw new InvalidOperationException(UiText.Get("文档或选区已改变，请重新定位。"));
        var token = Guid.NewGuid().ToString("N");
        var xml = SpreadsheetClipboardCodec.Write(region, formulas); var html = formulas ? null : HtmlClipboardCodec.Write(region);
        var text = formulas ? null : GridClipboard.Write(new(0, 0, region.Rows - 1, region.Columns - 1), (r, c) => CellInput.DisplayText(region.Cells[r * region.Columns + c].Value));
        await _clipboard.WriteAsync(new(text, token, xml, html));
        _visibleCapture = new(token, session, snapshot.Revision, region); _clipboardToken = null; _clipboardCells = null; _clipboardAxes = null;
        _status.Text = UiText.Format($"已复制 {region.Rows:N0} 个可见行、{region.Columns:N0} 列。可粘贴到匹配的可见目标范围。");
    }
    private async Task PasteVisibleAsync()
    {
        Document.Session.RequireWritable();
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var target = SelectedRange(); var id = snapshot.ActiveSheetId;
        void RequireCurrent()
        {
            if (session != Document.Session.SessionId || Document.Snapshot.Revision != snapshot.Revision || id != Document.Snapshot.ActiveSheetId ||
                target != SelectedRange() || !ReferenceEquals(snapshot.GetSheet(id), Document.Snapshot.GetSheet(id)))
                throw new InvalidOperationException(UiText.Get("文档或目标选区已改变，请重新粘贴。"));
        }
        var data = await _clipboard.ReadAsync(); RequireCurrent();
        ExternalClipboardRegion source;
        if (_visibleCapture is { } local && local.Token == data.Token)
        {
            if (local.Session != session || local.Revision != snapshot.Revision)
                throw new InvalidOperationException(UiText.Get("复制后文档或筛选已改变，请重新复制可见单元格。"));
            source = local.Data;
        }
        else source = data.SpreadsheetXml is { } xml ? SpreadsheetClipboardCodec.Read(xml, data.Biff8)
            : throw new NotSupportedException(UiText.Get("可见行粘贴需要类型化表格剪贴板，请先使用复制可见单元格或 Excel 表格复制。"));
        var rows = VisibleCellTransfer.Rows(snapshot, id, Rectangle(target));
        if (rows.Length != source.Rows || target.Columns != source.Columns)
            throw new ArgumentException(UiText.Format($"来源 {source.Rows:N0} 行 × {source.Columns:N0} 列，目标可见区域 {rows.Length:N0} 行 × {target.Columns:N0} 列。请调整目标选区。"));
        var dialog = new FieldsDialog(UiText.Get("粘贴到可见单元格"))
            .Note(UiText.Format($"将覆盖 {rows.Length:N0} 个可见行、{source.Columns:N0} 列，范围 {Rectangle(target)}。跳过 {target.Rows - rows.Length:N0} 个隐藏或筛除行。整批可一次撤销。"))
            .Choice("mode", UiText.Get("粘贴内容"), [UiText.Get("结果值和基本格式"), UiText.Get("只贴值"), UiText.Get("公式和值（保留目标格式）")]);
        if (!await dialog.ShowAsync(this)) return; RequireCurrent();
        var mode = dialog.Index("mode") switch { 1 => CellPasteMode.Values, 2 => CellPasteMode.Formulas, _ => CellPasteMode.All };
        if (await ExecuteAsync(new PasteVisibleCellsCommand(source, id, Rectangle(target), snapshot.Revision, mode), snapshot.Revision))
        { await CalculatePendingAsync(); _status.Text = UiText.Get("已粘贴到可见行，隐藏和筛除行保持原值。可一次撤销。"); }
    }
}
