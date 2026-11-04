using QuickExcel.Clipboard;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class PartialClipboardTests
{
    private static WorkbookSnapshot Book() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new(SheetId.New(), "Target")], styles: [new() { Italic = true }]);

    [Fact]
    public void PartialStylesPreserveSupportedPropertiesAndRetainDestinationForMissingStyles()
    {
        var book = Book(); using var session = new DocumentSession(book);
        var source = new ExternalClipboardRegion(1, 2, [new(CellValue.FromText("00123"), new() { Bold = true, Underline = "double", Indent = 2 }),
            new(CellValue.FromNumber(12))], limitations: ["unmodeled object"]);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), 0, allowPartial: true));
        var sheet = session.Snapshot.Sheets[0];
        var style = session.Snapshot.Styles[sheet.GetStyleId(0, 0)];
        Assert.True(style.Bold); Assert.Equal("double", style.Underline); Assert.Equal(2, style.Indent);
        Assert.Equal(0, sheet.GetStyleId(0, 1)); Assert.Equal(CellValue.FromNumber(12), sheet.GetCell(0, 1)!.Value);
        Assert.True(session.Undo()); Assert.Empty(session.Snapshot.Sheets[0].Cells); Assert.Equal(book.Styles, session.Snapshot.Styles);
        Assert.True(session.Redo()); Assert.Equal(CellValue.FromText("00123"), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void MissingResultsDoNotClearCellsOrReshapeMergesAndPartialPasteIsAtomic()
    {
        var book = Book(); using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.ActiveSheetId, 0, 1), CellContent.FromLiteral(CellValue.FromText("keep")))]));
        var before = session.Snapshot;
        var source = new ExternalClipboardRegion(1, 2, [new(CellValue.FromNumber(5), new()),
            new(CellValue.Blank, new(), Formula: "=1+2", MissingResult: true)], merges: [new(0, 0, 0, 1)]);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), before.Revision, allowPartial: true));
        Assert.Equal(CellValue.FromNumber(5), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText("keep"), session.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Empty(session.Snapshot.Sheets[0].MergedCells);
        Assert.True(session.Undo()); Assert.Equal(before.Sheets[0].Cells, session.Snapshot.Sheets[0].Cells);
    }

    [Fact]
    public void AllMissingResultsAreANoOpWithoutAnUndoEntry()
    {
        var book = Book(); using var session = new DocumentSession(book);
        var source = new ExternalClipboardRegion(1, 1, [new(CellValue.Blank, new() { Bold = true }, MissingResult: true)]);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), 0, allowPartial: true));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void CellPastePreservesExistingFloatingObjectAndItsAnchor()
    {
        var id = SheetId.New(); var image = new SheetImage([1, 2, 3], "png", 0, 0, 100, 100);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(id, "Target", images: [image])]);
        using var session = new DocumentSession(book);
        var source = new ExternalClipboardRegion(1, 1, [new(CellValue.FromNumber(42), new())]);
        session.Execute(new PasteExternalCellsCommand(source, id, new(0, 0, 0, 0), 0, allowPartial: true));
        Assert.Equal(CellValue.FromNumber(42), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(image, Assert.Single(session.Snapshot.Sheets[0].Images));
        Assert.True(session.Undo()); Assert.Equal(image, Assert.Single(session.Snapshot.Sheets[0].Images));
    }

    [Fact]
    public void UnsupportedFormulasBecomeCachedValuesWhileSupportedFormulasRemainFormulas()
    {
        var book = Book(); using var session = new DocumentSession(book);
        var source = new ExternalClipboardRegion(1, 3, [new(CellValue.FromNumber(3), Formula: "=1+2"),
            new(CellValue.FromNumber(7), FormulaRejection: "external reference"),
            new(CellValue.FromText("=literal"), Formula: "=UnknownFunction()")]);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), 0, CellPasteMode.Formulas, allowPartial: true));
        var sheet = session.Snapshot.Sheets[0];
        Assert.Equal("=1+2", sheet.GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal(CellContentKind.Literal, sheet.GetCell(0, 1)!.Content.Kind);
        Assert.Equal(CellValue.FromNumber(7), sheet.GetCell(0, 1)!.Value);
        Assert.Equal(CellValue.FromText("=literal"), sheet.GetCell(0, 2)!.Value);
    }

    [Fact]
    public void HtmlPartialPasteKeepsTextLiteralAndSupportedCellFormats()
    {
        const string html = "<table><tr><td style='font-weight:bold'><a href='https://example.invalid'>=1+2</a><img src='file:///private'/></td></tr></table>";
        var source = HtmlClipboardCodec.Read(HtmlClipboardCodec.Wrap(html, 0, html.Length));
        var plan = ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All);
        var book = Book(); using var session = new DocumentSession(book);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), 0, plan.Mode, allowPartial: true));
        Assert.Equal(CellValue.FromText("=1+2"), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.True(session.Snapshot.Styles[session.Snapshot.Sheets[0].GetStyleId(0, 0)].Bold);
        Assert.Contains("显示文本", plan.Notice); Assert.Contains("支持的部分", plan.Notice);
    }

    [Fact]
    public void HtmlTextFallbackOnlyUsesSelectedFragmentAndDoesNotExposeScriptOrContext()
    {
        const string before = "<html><body><table><tr><td>unselected</td><td>";
        const string selected = "<div>chosen<br>=1+2<script>secret()</script><img alt='picture' src='https://example.invalid/image'></div>";
        const string after = "</td><td>also unselected</td></tr></table></body></html>";
        var bytes = HtmlClipboardCodec.Wrap(before + selected + after, before.Length, before.Length + selected.Length);
        Assert.Throws<NotSupportedException>(() => HtmlClipboardCodec.Read(bytes));
        Assert.Equal("chosen\n=1+2picture", HtmlClipboardCodec.ReadDisplayText(bytes));
    }
}
