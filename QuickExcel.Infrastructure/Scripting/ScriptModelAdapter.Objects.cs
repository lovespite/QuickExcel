using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Scripting;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private static readonly JsonElement ApiEnums = JsonDocument.Parse(ScriptCapabilities.Current.Json).RootElement.GetProperty("enums").Clone();
    private static int EnumNumber(string type, string name) => ApiEnums.GetProperty(type).GetProperty(name).GetInt32();
    private static string EnumName(string type, JsonElement value) => ApiEnums.GetProperty(type).EnumerateObject().FirstOrDefault(p => p.Value.GetInt32() == Int(value)).Name ?? throw Invalid("无效枚举值。");
    private static void Fields(JsonElement value, params string[] names)
    {
        if (value.ValueKind != JsonValueKind.Object) throw Invalid("需要选项对象。");
        if (value.EnumerateObject().Any(p => !names.Contains(p.Name))) throw Unsupported("选项含尚未支持的字段。");
    }
    private static bool Overlap(CellRectangle a, CellRectangle b) => a.Top <= b.Bottom && b.Top <= a.Bottom && a.Left <= b.Right && b.Left <= a.Right;
    private static int RuleCost(IEnumerable<CellRectangle> ranges)
    {
        var count = ranges.Sum(r => r.Count);
        if (count > MaxWriteCells) throw new ModelFailure("BudgetExceeded", "对象规则修改超过 20000 格。");
        return checked((int)count);
    }
    private void WriteRules(WorksheetSnapshot sheet, IEnumerable<CellRectangle> ranges, DocumentCommand command)
    {
        var cost = RuleCost(ranges); Spend(Math.Max(1, cost), true);
        Model.ExecuteObjectCommand(sheet.Id, Math.Max(1, cost), command);
    }
    private static string Reference(WorksheetSnapshot sheet, CellRectangle range) => FormulaSyntax.Print(new ReferenceNode(sheet.Name,
        range.Top, range.Left, range.Bottom, range.Right, true, true, true, true, IsRange: range.Count != 1));
    private object CommentHandle(WorksheetSnapshot sheet, WorksheetComment c) => Handle(c.ParentId is null ? "Comment" : "CommentReply", new { sheetId = sheet.Id.ToString(), commentId = c.Id.ToString() });
    private object? Comments(WorksheetSnapshot? target, string member, JsonElement[] args)
    {
        var sheets = target is null ? Model.View.Sheets : [target];
        var entries = sheets.SelectMany(s => s.Data.ThreadComments.Select(c => (Sheet: s, Comment: c))).ToArray();
        Spend(Math.Max(1, entries.Length), false);
        if (member == "getComments") { Arity(args, 0); return entries.Where(e => e.Comment.ParentId is null).Select(e => CommentHandle(e.Sheet, e.Comment)).ToArray(); }
        if (member is "getComment" or "getCommentByReplyId")
        {
            Arity(args, 1); var key = Text(args[0]); var entry = entries.FirstOrDefault(e => e.Comment.Id.ToString() == key);
            if (entry.Comment is null) { if (member == "getComment") return null; throw new ModelFailure("ItemNotFound", "找不到评论回复。"); }
            if (member == "getCommentByReplyId")
            { if (entry.Comment.ParentId is not { } parent) throw Invalid("需要回复 ID。"); entry = entries.Single(e => e.Comment.Id == parent); }
            else if (entry.Comment.ParentId is not null) return null;
            return CommentHandle(entry.Sheet, entry.Comment);
        }
        Arity(args, 1, member == "addComment" ? 3 : 1);
        var sheet = target ?? (args[0].ValueKind == JsonValueKind.Object ? Model.View.GetSheet(new(Guid.Parse(Text(args[0].GetProperty("sheetId"))))) : Model.View.GetSheet(Model.View.ActiveSheetId));
        var range = Address(args[0], sheet.Id); if (range.Count != 1) throw Invalid("评论需要单个单元格。");
        var existing = sheet.Data.ThreadComments.FirstOrDefault(c => c.Position == new CellPosition(range.Top, range.Left) && c.ParentId is null);
        if (member == "getCommentByCell") return existing is null ? throw new ModelFailure("ItemNotFound", "单元格没有线程评论。") : CommentHandle(sheet, existing);
        if (member != "addComment" || args.Length < 2) throw Invalid("需要评论内容。");
        if (existing is not null) throw Invalid("单元格已有线程评论。");
        PlainComment(args, 1);
        var comment = new WorksheetComment(Guid.NewGuid(), new(range.Top, range.Left), Text(args[1]), "QuickExcel", DateTimeOffset.UtcNow, LocalCommentAuthor);
        Write(sheet.Id, range, new SetScriptCommentsCommand(sheet.Id, sheet.Data.ThreadComments.Add(comment)));
        return CommentHandle(sheet, comment);
    }
    private static readonly Guid LocalCommentAuthor = new("716c0b1e-7739-4ca6-9daa-e9e2c4cab311");
    private static void PlainComment(JsonElement[] args, int content)
    {
        if (args[content].ValueKind != JsonValueKind.String || args.Length > content + 1 && EnumName("ContentType", args[content + 1]) != "plain") throw Unsupported("线程评论支持纯文本，不解析提及、不发送通知。");
        WorkbookLimits.ValidateText(Text(args[content]));
    }
    private object? CommentApi(WorksheetSnapshot sheet, JsonElement data, string member, JsonElement[] args)
    {
        var key = Text(data.GetProperty("commentId"));
        var current = sheet.Data.ThreadComments.FirstOrDefault(c => c.Id.ToString() == key) ?? throw new ModelFailure("InvalidReference", "评论句柄已失效。");
        var root = current.ParentId is { } parent ? sheet.Data.ThreadComments.Single(c => c.Id == parent) : current;
        var range = new CellRectangle(current.Position.Row, current.Position.Column, current.Position.Row, current.Position.Column);
        if (member.StartsWith("get", StringComparison.Ordinal))
        {
            Arity(args, member == "getCommentReply" ? 1 : 0); Spend(1, false);
            return member switch
            {
                "getId" => key, "getAuthorName" => current.Author, "getContent" or "getRichContent" => current.Content,
                "getContentType" => EnumNumber("ContentType", "plain"), "getCreationDate" => current.Created.ToString("O"),
                "getResolved" => root.Resolved, "getLocation" => RangeHandle(sheet.Id, range), "getParentComment" => CommentHandle(sheet, root),
                "getReplies" => sheet.Data.ThreadComments.Where(c => c.ParentId == current.Id).Select(c => CommentHandle(sheet, c)).ToArray(),
                "getCommentReply" => sheet.Data.ThreadComments.FirstOrDefault(c => c.ParentId == current.Id && c.Id.ToString() == Text(args[0])) is { } reply ? CommentHandle(sheet, reply) : null,
                _ => throw Unsupported(member)
            };
        }
        var next = sheet.Data.ThreadComments;
        if (member == "delete") { Arity(args, 0); next = next.Where(c => c.Id != current.Id && c.ParentId != current.Id).ToImmutableArray(); }
        else if (member == "setContent") { Arity(args, 1); PlainComment(args, 0); next = next.Replace(current, current with { Content = Text(args[0]) }); }
        else if (member == "setResolved") { Arity(args, 1); next = next.Replace(root, root with { Resolved = Bool(args[0]) }); }
        else if (member == "addCommentReply")
        {
            Arity(args, 1, 2); PlainComment(args, 0);
            if (root.Resolved) throw Invalid("请先重新打开已解决的评论。");
            var reply = new WorksheetComment(Guid.NewGuid(), current.Position, Text(args[0]), "QuickExcel", DateTimeOffset.UtcNow, LocalCommentAuthor, root.Id);
            Write(sheet.Id, range, new SetScriptCommentsCommand(sheet.Id, next.Add(reply))); return CommentHandle(sheet, reply);
        }
        else throw Unsupported(member);
        Write(sheet.Id, range, new SetScriptCommentsCommand(sheet.Id, next)); return null;
    }
}
