using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Local plain-text discussions; no account lookup or notifications.</summary>
public sealed record WorksheetComment(Guid Id, CellPosition Position, string Content, string Author,
    DateTimeOffset Created, Guid PersonId, Guid? ParentId = null, bool Resolved = false,
    string? UserId = null, string? ProviderId = null);

public sealed class SetScriptCommentsCommand(SheetId sheetId, IEnumerable<WorksheetComment> comments) : DocumentCommand
{
    private readonly ImmutableArray<WorksheetComment> _comments = comments.Take(4097).ToImmutableArray();
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Data.ThreadComments.SequenceEqual(_comments)) return new(snapshot, [], false);
        if (sheet.Data.CommentFallback || snapshot.FileState?.CompatibilityDiagnostics.Any(d => d.Code == "xlsx.unmodeled" && d.Feature.Contains("threadedComment", StringComparison.OrdinalIgnoreCase)) == true)
            throw new NotSupportedException("含旧版显示副本或未建模扩展的线程评论只支持原位保留，不能编辑或扁平化。");
        Validate(_comments);
        if (_comments.Any(c => sheet.Annotations.Notes.ContainsKey(c.Position))) throw new NotSupportedException("目标含旧版批注，不能替换为线程评论。");
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(s.Data with { ThreadComments = _comments }) : s)), [], false, CalculationUnchanged: true);
    }
    public static void Validate(ImmutableArray<WorksheetComment> comments)
    {
        if (comments.IsDefault || comments.Length > 4096 || comments.Select(c => c.Id).Distinct().Count() != comments.Length || comments.Sum(c => (long)c.Content.Length * 2) > 4 * 1024 * 1024)
            throw new ArgumentException("线程评论数量或容量超限。");
        var ids = comments.ToDictionary(c => c.Id);
        foreach (var c in comments)
        {
            WorkbookLimits.ValidatePosition(c.Position.Row, c.Position.Column); WorkbookLimits.ValidateText(c.Content);
            if (c.Id == Guid.Empty || c.PersonId == Guid.Empty || c.Author.Length > 1024 || c.UserId?.Length > 1024 || c.ProviderId?.Length > 1024 ||
                c.ParentId is { } p && (!ids.TryGetValue(p, out var parent) || parent.ParentId is not null || parent.Position != c.Position))
                throw new ArgumentException("线程评论身份或父级无效。");
        }
        if (comments.Where(c => c.ParentId is null).GroupBy(c => c.Position).Any(g => g.Count() != 1)) throw new ArgumentException("单元格已有线程评论。");
    }
}

/// <summary>Do not remove an overlapping validation rule outside the requested range.</summary>
public sealed class SetScriptValidationCommand(SheetId sheetId, CellRectangle range, WorksheetValidation? rule) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        bool Inside(CellRectangle r) => range.Contains(r.Top, r.Left) && range.Contains(r.Bottom, r.Right);
        var overlaps = sheet.Data.Validations.Where(v => v.Ranges.Any(r => ValidationSafety.Overlaps(r, range))).ToArray();
        if (overlaps.Any(v => v.Ranges.Any(r => !Inside(r)))) throw new NotSupportedException("目标只覆盖现有验证规则的一部分，请使用完整规则范围。");
        rule?.ValidateBinding(snapshot, sheetId);
        var remaining = sheet.Data.Validations.Except(overlaps).ToImmutableArray();
        if (rule is not null) remaining = remaining.Add(rule);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(s.Data with { Validations = remaining }) : s)), [], false, CalculationUnchanged: true);
    }
}
