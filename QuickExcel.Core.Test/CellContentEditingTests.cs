using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class CellContentEditingTests
{
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void LinksPreserveRichTextNotesAndFormulaWhileRoundtripping(string extension)
    {
        var rich = new RichCellText([new("中文", new(Bold: true)), new(" link")]);
        var sheet = Sheet("S", (0, 0, new(CellContent.FromRichText(rich))), (0, 1, F("=2+2")));
        using var session = new DocumentSession(Book(sheet));
        var note = new CellNote("A", new([new TextRun("note")]));
        session.Execute(new SetNoteCommand(sheet.Id, new(0, 0), note));
        var link = new CellHyperlink("https://example.com/a?q=%E4%B8%AD", Tooltip: "提示");
        session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 0), link));
        Assert.Equal(rich, session.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Content.RichText);
        Assert.True(session.Undo()); Assert.Empty(session.Snapshot.GetSheet(sheet.Id).Annotations.Links); Assert.True(session.Redo());
        session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 1), new(null, "'S'!A1")));
        session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 2), new("mailto:person@example.com")));
        var folder = Path.Combine(AppContext.BaseDirectory, "content", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        var file = Path.Combine(folder, "links." + extension); var storage = new TypedWorkbookStorage(); storage.Save(file, session.Snapshot);
        var reopened = storage.Open(file).Sheets[0];
        Assert.Equal(link, reopened.Annotations.Links[new(0, 0)]); Assert.Equal(note, reopened.Annotations.Notes[new(0, 0)]);
        Assert.Equal(rich, reopened.GetCell(0, 0)!.Content.RichText); Assert.Equal("=2+2", reopened.GetCell(0, 1)!.Content.FormulaText);
        Assert.Equal("'S'!A1", reopened.Annotations.Links[new(0, 1)].Location); Assert.Equal("mailto:person@example.com", reopened.Annotations.Links[new(0, 2)].Target);
        session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 0), null)); Assert.False(session.Snapshot.GetSheet(sheet.Id).Annotations.Links.ContainsKey(new(0, 0)));
    }
    [Theory]
    [InlineData("file:///C:/test.exe", null)]
    [InlineData("javascript:alert(1)", null)]
    [InlineData(null, "Missing!A1")]
    [InlineData(null, "SUM(A1:A2)")]
    public void UnsupportedTargetsFailAtomically(string? target, string? location)
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        Assert.Throws<ArgumentException>(() => session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 0), new(target, location))));
        Assert.False(session.CanUndo); Assert.Equal(0, session.Snapshot.Revision);
    }
    [Fact]
    public void HyperlinkRestrictionIsIndependentOfObjectProtection()
    {
        var sheet = Sheet("S").WithData(new() { Protection = new() { Objects = false, InsertHyperlinks = true } });
        using var session = new DocumentSession(Book(sheet));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 0), new("https://example.com"))));
        Assert.False(session.CanUndo);
    }
}
