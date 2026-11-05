using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class QuickCellStyleTests
{
    [Fact]
    public void PresetsPreserveValuesAndUndoOneRangeAsOneEntry()
    {
        var sheet = Sheet("S", (0, 0, N(1.25)), (1, 0, F("=1+2")));
        using var session = new DocumentSession(Book(sheet));
        foreach (var preset in QuickCellStyles.BuiltIn)
        {
            var before = session.Snapshot;
            Assert.False(string.IsNullOrWhiteSpace(QuickCellStyles.DisplayName(preset)));
            session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0), new(sheet.Id, 1, 0)], preset.Patch));
            Assert.Equal(CellValue.FromNumber(1.25), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
            Assert.Equal("=1+2", session.Snapshot.Sheets[0].GetCell(1, 0)!.Content.FormulaText);
            if (session.Snapshot.Revision == before.Revision) continue;
            Assert.True(session.Undo()); Assert.Equal(before.Sheets[0].GetStyleId(0, 0), session.Snapshot.Sheets[0].GetStyleId(0, 0));
            Assert.False(session.CanUndo); Assert.True(session.Redo()); Assert.True(session.Undo());
        }
    }

    [Fact]
    public void OnlyCheckedCategoriesChangeAndCanExplicitlyClearValues()
    {
        var before = new StyleDefinition { NumberFormatId = 4, Bold = true, Italic = true, Underline = "double", FillColor = "FFFF0000", FontColor = "FF008800", Locked = false, Horizontal = 3, WrapText = true, LeftBorder = 5, TextRotation = 90 };
        var template = new QuickCellStyle("test", "test", "test", new(), CellStyleParts.Fill);
        Assert.Equal(before with { FillColor = null }, QuickCellStyle.EditFormat(before, template.Patch));
        template = template with { Parts = CellStyleParts.Number };
        Assert.Equal(before with { NumberFormatId = 0 }, QuickCellStyle.EditFormat(before, template.Patch));
        template = template with { Parts = CellStyleParts.Alignment | CellStyleParts.Protection };
        Assert.Equal(before with { Horizontal = 0, WrapText = false, Locked = true, TextRotation = 0 }, QuickCellStyle.EditFormat(before, template.Patch));
        template = template with { Parts = CellStyleParts.Font };
        var after = QuickCellStyle.EditFormat(before, template.Patch);
        Assert.Null(after.FontColor); Assert.Null(after.Underline); Assert.False(after.Bold); Assert.False(after.Italic);
        Assert.Equal(before.FillColor, after.FillColor); Assert.Equal(before.NumberFormatId, after.NumberFormatId); Assert.Equal(before.LeftBorder, after.LeftBorder);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void AppliedStylesRoundTripIndependentOfLocalLibrary(string extension)
    {
        var sheet = Sheet("S", (0, 0, N(42)));
        using var session = new DocumentSession(Book(sheet));
        var style = new QuickCellStyle(Guid.NewGuid().ToString("N"), "本机样式", "自定义", new() { FillColor = "FF336699", FontColor = "FFFFFFFF", Bold = true, NumberFormatCode = "0.000", BottomBorder = 6, Locked = false }, CellStyleParts.All);
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], style.Patch));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            var format = reopened.Styles[reopened.Sheets[0].GetStyleId(0, 0)];
            Assert.Equal("FF336699", format.FillColor); Assert.True(format.Bold); Assert.Equal(6, format.BottomBorder);
            Assert.Equal("0.000", format.NumberFormatCode); Assert.False(format.Locked); Assert.Equal(CellValue.FromNumber(42), reopened.Sheets[0].GetCell(0, 0)!.Value);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void FontStylesPreserveRichTextAndNonFontStylesLeaveRunsUntouched()
    {
        var rich = new RichCellText([new TextRun("one", new(Name: "Arial", Size: 14, Italic: true)), new TextRun("two")]);
        var id = SheetId.New(); var sheet = new WorksheetSnapshot(id, "S", new Dictionary<CellPosition, CellSnapshot> { [new(0, 0)] = new(CellContent.FromRichText(rich)) });
        using var session = new DocumentSession(Book(sheet));
        var fill = new QuickCellStyle("fill", "fill", "fill", new() { FillColor = "FFFFFF00" }, CellStyleParts.Fill);
        session.Execute(new SetCellStyleCommand([new(id, 0, 0)], fill.Patch));
        Assert.Equal(rich, session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
        session.Execute(new SetCellStyleCommand([new(id, 0, 0)], QuickCellStyles.BuiltIn.Single(s => s.Id == "good").Patch));
        var after = session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText!;
        Assert.Equal(2, after.Runs.Length); Assert.Equal(rich.Text, after.Text); Assert.All(after.Runs, r => Assert.Equal("FF006100", r.Font!.Color));
    }

    [Fact]
    public void LibraryReloadsRejectsDuplicatesAndNeverOverwritesConcurrentOrCorruptFiles()
    {
        var directory = Path.Combine(Path.GetTempPath(), "quick-styles-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "styles.json");
        try
        {
            var first = new QuickCellStyleStore(path); first.Load();
            var style = new QuickCellStyle(Guid.NewGuid().ToString("N"), "My Style", "自定义", new() { Bold = true }, CellStyleParts.Font);
            first.Save([style]); var second = new QuickCellStyleStore(path); second.Load(); Assert.Equal(style, Assert.Single(second.Styles));
            Assert.Throws<ArgumentException>(() => first.Save([style, style with { Id = Guid.NewGuid().ToString("N"), Name = "MY STYLE" }]));
            Assert.Throws<ArgumentException>(() => first.Save([style with { Parts = CellStyleParts.None }]));
            first.Save([style with { Name = "Updated" }]); Assert.Throws<InvalidOperationException>(() => second.Save([]));
            second.Load(); Assert.Equal("Updated", Assert.Single(second.Styles).Name);
            File.WriteAllText(path, "{broken"); second.Load(); Assert.NotNull(second.Error);
            Assert.Throws<InvalidOperationException>(() => second.Save([])); Assert.Equal("{broken", File.ReadAllText(path));
        }
        finally { Directory.Delete(directory, true); }
    }
}
