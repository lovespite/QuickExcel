using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RichTextAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var sheet = snapshot.Sheets[SheetIndex]; var position = _grid.Current;
        var cell = sheet.GetCell(position.Row, position.Column);
        if (cell?.SpillOwner is not null || cell?.Content.Kind == CellContentKind.Formula || cell?.Value.Kind is not (null or CellValueKind.Blank or CellValueKind.Text))
            throw new InvalidOperationException(UiText.Get("富文本编辑仅适用于文字或空白单元格。"));
        var source = cell?.Content.RichText ?? new RichCellText([new(cell?.Value is { Kind: CellValueKind.Text } text ? text.TextValue : "")]);
        var dialog = new RichTextDialog(source, snapshot.Styles[sheet.GetStyleId(position.Row, position.Column)]);
        if (!await ShowOwnedAsync<bool>(dialog)) return;
        RequireContentDialogCurrent(session, snapshot);
        if (!source.Equals(dialog.Draft)) await ExecuteAsync(new SetCellsCommand([new(new(sheet.Id, position.Row, position.Column), CellContent.FromRichText(dialog.Draft, cell?.Content.StyleId ?? 0))]), snapshot.Revision);
    }
    private void RequireContentDialogCurrent(Guid session, WorkbookSnapshot snapshot)
    {
        if (session != Document.Session.SessionId || snapshot.Revision != Document.Snapshot.Revision || snapshot.ActiveSheetId != Document.Snapshot.ActiveSheetId)
            throw new InvalidOperationException(UiText.Get("文档已改变，请重新打开内容编辑窗口。"));
    }
    private async Task HyperlinkAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId;
        var sheet = snapshot.Sheets[SheetIndex]; var position = _grid.Current; var old = sheet.Annotations.Links.GetValueOrDefault(position);
        var dialog = new FieldsDialog(UiText.Get("超链接"))
            .Note(UiText.Get("链接仅保存和查看，不自动打开。显示名称不会改写单元格文字；内部位置示例：'Sheet1'!A1。"))
            .Text("target", UiText.Get("网页或邮件地址"), old?.Target ?? "")
            .Text("location", UiText.Get("工作簿内位置"), old?.Location ?? "")
            .Text("display", UiText.Get("显示名称"), old?.Display ?? "")
            .Text("tooltip", UiText.Get("提示文字"), old?.Tooltip ?? "")
            .Check("remove", UiText.Get("删除此链接"));
        if (!await dialog.ShowAsync(this)) return;
        RequireContentDialogCurrent(session, snapshot);
        string? Field(string key) => string.IsNullOrWhiteSpace(dialog.String(key)) ? null : dialog.String(key);
        await ExecuteAsync(new SetHyperlinkCommand(sheet.Id, position, dialog.Boolean("remove") ? null : new(Field("target"), Field("location"), Field("display"), Field("tooltip"))), snapshot.Revision);
    }
    private async Task ThreadCommentsAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId;
        var sheet = snapshot.Sheets[SheetIndex]; var position = _grid.Current;
        var comments = sheet.Data.ThreadComments.Where(c => c.Position == position).OrderBy(c => c.ParentId is not null).ThenBy(c => c.Created).ToArray();
        var root = comments.FirstOrDefault(c => c.ParentId is null);
        var dialog = new FieldsDialog(UiText.Get("线程评论"));
        var summary = string.Join("\n\n", comments.Select(c => $"{c.Author} · {c.Created.ToLocalTime():yyyy-MM-dd HH:mm} · {UiText.Get(c.Resolved ? "已解决" : "未解决")}\n{c.Content}"));
        dialog.Text("history", UiText.Get("讨论记录"), summary, true);
        ((TextBox)dialog.Input("history")).IsReadOnly = true;
        var preserved = sheet.Data.CommentFallback || snapshot.FileState?.PreservedAttachments.Any(a => a.Feature == ExtendedObjectSafety.ThreadedComments && a.SheetName == sheet.Name) == true ||
            snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("threadedComment", StringComparison.OrdinalIgnoreCase)) == true;
        if (Document.IsReadOnly || preserved)
        {
            dialog.Note(UiText.Get("此讨论仅供查看，保留原始评论及旧版显示副本。")); await dialog.ShowAsync(this); return;
        }
        dialog.Note(UiText.Get("评论存储在工作簿内，不发送通知。删除主评论会同时删除其回复。"));
        if (root is not null)
        {
            dialog.Choice("selected", UiText.Get("选择评论或回复"), comments.Select(c => c.Author + ": " + (c.Content.Length > 80 ? c.Content[..80] + "…" : c.Content)).ToArray());
            dialog.Choice("operation", UiText.Get("操作"), [UiText.Get("添加回复"), UiText.Get("解决讨论"), UiText.Get("重新打开讨论"), UiText.Get("删除选中评论"), UiText.Get("修改选中评论")]);
        }
        dialog.Text("author", UiText.Get("作者"), "QuickExcel").Text("text", UiText.Get("评论内容"), "", true);
        if (!await dialog.ShowAsync(this)) return;
        RequireContentDialogCurrent(session, snapshot);
        var items = sheet.Data.ThreadComments.ToList(); var operation = root is null ? 0 : dialog.Index("operation");
        var selected = root is null ? null : comments[dialog.Index("selected")];
        if (operation is 0 or 4 && string.IsNullOrWhiteSpace(dialog.String("text"))) throw new ArgumentException(UiText.Get("请输入评论内容。"));
        switch (operation)
        {
            case 0:
                if (root?.Resolved == true) throw new InvalidOperationException(UiText.Get("请先重新打开已解决的讨论，再添加回复。"));
                var author = dialog.String("author").Trim(); if (author.Length is 0 or > 1024) throw new ArgumentException(UiText.Get("请输入有效的作者名称。"));
                var person = snapshot.Sheets.SelectMany(s => s.Data.ThreadComments).FirstOrDefault(c => c.Author == author && c.UserId is null && c.ProviderId is null)?.PersonId ?? Guid.NewGuid();
                items.Add(new(Guid.NewGuid(), position, dialog.String("text"), author, DateTimeOffset.UtcNow, person, root?.Id)); break;
            case 1: case 2:
                items[items.FindIndex(c => c.Id == root!.Id)] = root! with { Resolved = operation == 1 }; break;
            case 3:
                items.RemoveAll(c => c.Id == selected!.Id || c.ParentId == selected.Id); break;
            case 4:
                items[items.FindIndex(c => c.Id == selected!.Id)] = selected! with { Content = dialog.String("text") }; break;
        }
        await ExecuteAsync(new SetScriptCommentsCommand(sheet.Id, items), snapshot.Revision);
    }
}
