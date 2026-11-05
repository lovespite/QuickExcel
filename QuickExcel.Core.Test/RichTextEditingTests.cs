using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class RichTextEditingTests
{
    [Fact]
    public void InsertReplaceMergeAndDeletePreserveFontsAndUnicode()
    {
        var red = new TextRunFont(Color: "FFFF0000"); var source = new RichCellText([new("中文", red), new("😀tail")]);
        var inserted = RichTextEditing.Replace(source, 2, 0, "\n新");
        Assert.Equal("中文\n新😀tail", inserted.Text); Assert.Equal("中文\n新", inserted.Runs[0].Text); Assert.Equal("FFFF0000", inserted.Runs[0].Font!.Color);
        var replaced = RichTextEditing.Replace(source, 1, 3, "替换"); Assert.Equal("中替换tail", replaced.Text); Assert.Equal(2, replaced.Runs.Length);
        Assert.Throws<ArgumentException>(() => RichTextEditing.Replace(source, 3, 0, "bad"));
        Assert.Throws<ArgumentException>(() => RichTextEditing.Format(source, 2, 1, new(Bold: true), new()));
        Assert.Equal("", RichTextEditing.Replace(source, 0, source.Text.Length, "").Text);
    }
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void PartialFormattingPreservesInheritedFontAttachmentsAndUndo(string extension)
    {
        var source = new RichCellText([new("中文 text")]); var style = new StyleDefinition { FontName = "Arial", FontSize = 18, FontColor = "FF123456" };
        var formatted = RichTextEditing.Format(source, 0, 2, new(Bold: true), style);
        Assert.Equal(2, formatted.Runs.Length); Assert.True(formatted.Runs[0].Font!.Bold); Assert.Equal("Arial", formatted.Runs[0].Font!.Name); Assert.Null(formatted.Runs[1].Font);
        var sheet = Sheet("S", (0, 0, new(CellContent.FromRichText(source))));
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [style]));
        session.Execute(new SetHyperlinkCommand(sheet.Id, new(0, 0), new("https://example.com")));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromRichText(formatted))]));
        Assert.True(session.Undo()); Assert.Equal(source, session.Snapshot.GetSheet(sheet.Id).GetCell(0, 0)!.Content.RichText); Assert.True(session.Redo());
        var folder = Path.Combine(AppContext.BaseDirectory, "rich-edit", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder); var file = Path.Combine(folder, "text." + extension);
        var storage = new TypedWorkbookStorage(); storage.Save(file, session.Snapshot); var reopened = storage.Open(file).Sheets[0];
        Assert.Equal(formatted.Text, reopened.GetCell(0, 0)!.Content.RichText!.Text); Assert.True(reopened.GetCell(0, 0)!.Content.RichText!.Runs[0].Font!.Bold);
        Assert.Equal("https://example.com", reopened.Annotations.Links[new(0, 0)].Target);
    }
    [Fact]
    public void ColorEditsReplaceThemeAndTintWithoutChangingUnselectedRuns()
    {
        var source = new RichCellText([new("abc", new(Color: "FF112233", ThemeColor: 2, Tint: .2, Underline: "single"))]);
        var result = RichTextEditing.Format(source, 1, 1, new(Color: "FFABCDEF"), new());
        Assert.Equal(3, result.Runs.Length); Assert.Equal(2, result.Runs[0].Font!.ThemeColor); Assert.Null(result.Runs[1].Font!.ThemeColor); Assert.Null(result.Runs[1].Font!.Tint);
        Assert.Equal("FFABCDEF", result.Runs[1].Font!.RawRgb); Assert.Equal("single", result.Runs[1].Font!.Underline);
        Assert.Throws<ArgumentException>(() => RichTextEditing.Format(source, 0, 0, new(Bold: true), new()));
        Assert.Throws<ArgumentException>(() => RichTextEditing.Replace(source, 0, 0, new string('x', WorkbookLimits.MaxCellTextLength)));
    }
}
