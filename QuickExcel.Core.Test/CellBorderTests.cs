using System.IO.Compression;
using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class CellBorderTests
{
    private static CellAddress[] Area(SheetId sheet, int top = 1, int left = 1, int bottom = 3, int right = 3) =>
        Enumerable.Range(top, bottom - top + 1).SelectMany(r => Enumerable.Range(left, right - left + 1).Select(c => new CellAddress(sheet, r, c))).ToArray();
    private static StyleDefinition Style(WorkbookSnapshot book, int row, int column) => book.Styles[book.Sheets[0].GetStyleId(row, column)];

    [Fact]
    public void OutsideAndInsideUseSelectionEdgesAndOneUndoEntry()
    {
        var sheet = Sheet("S", (1, 1, N(7)), (2, 2, F("=B2*2"))); using var session = new DocumentSession(Book(sheet));
        session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())); var before = session.Snapshot;
        session.Execute(new SetCellBordersCommand(Area(sheet.Id), CellBorders.Preset(BorderPreset.Outside, 10, "FF0070C0")));
        Assert.Equal(10, Style(session.Snapshot, 1, 1).LeftBorder); Assert.Equal(10, Style(session.Snapshot, 3, 3).BottomBorder);
        Assert.Equal(0, Style(session.Snapshot, 2, 2).LeftBorder); Assert.Equal(0, Style(session.Snapshot, 1, 1).RightBorder);
        Assert.Equal(CalculationState.Fresh, session.Snapshot.Sheets[0].GetCell(2, 2)!.Calculation!.State);
        Assert.Equal(CellValue.FromNumber(14), session.Snapshot.Sheets[0].GetCell(2, 2)!.Value);
        Assert.True(session.Undo()); Assert.Equal(before.Sheets[0].Cells.Select(p => (p.Key, p.Value.Content)).OrderBy(p => p.Key.Row), session.Snapshot.Sheets[0].Cells.Select(p => (p.Key, p.Value.Content)).OrderBy(p => p.Key.Row)); Assert.False(session.CanUndo);
        session.Redo(); session.Execute(new SetCellBordersCommand(Area(sheet.Id), CellBorders.Preset(BorderPreset.Inside, 4)));
        Assert.Equal(4, Style(session.Snapshot, 2, 2).LeftBorder); Assert.Equal(4, Style(session.Snapshot, 2, 1).RightBorder);
        Assert.Equal(10, Style(session.Snapshot, 2, 1).LeftBorder); Assert.Equal("FF0070C0", Style(session.Snapshot, 2, 1).LeftBorderColor);
    }

    [Fact]
    public void ClearRemovesOpposingNeighborAndDiagonalsButKeepsRichTextAndUnrelatedEdges()
    {
        var rich = new RichCellText([new("rich", new(Bold: true)), new(" text")]);
        var sheet = Sheet("S", (1, 1, new(CellContent.FromRichText(rich)))); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 1, 2)], new(LeftBorder: 5, BottomBorder: 3)));
        session.Execute(new SetCellBordersCommand([new(sheet.Id, 1, 1)], new(Right: new(6, "FFFF0000"), Diagonal: new(9), DiagonalUp: true, DiagonalDown: true)));
        Assert.Equal(0, Style(session.Snapshot, 1, 2).LeftBorder); Assert.Equal(3, Style(session.Snapshot, 1, 2).BottomBorder);
        session.Execute(new SetCellBordersCommand([new(sheet.Id, 1, 1)], CellBorders.Preset(BorderPreset.None)));
        Assert.False(Style(session.Snapshot, 1, 1).DiagonalUp); Assert.Equal(0, Style(session.Snapshot, 1, 1).RightBorder);
        Assert.Equal(rich, session.Snapshot.Sheets[0].GetCell(1, 1)!.Content.RichText);
        session.Undo(); Assert.True(Style(session.Snapshot, 1, 1).DiagonalDown);
    }

    [Fact]
    public void AutomaticColorResetsOldColorAndNoncontiguousSelectionDoesNotFillTheGap()
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet)); var addresses = new[] { new CellAddress(sheet.Id, 0, 0), new CellAddress(sheet.Id, 2, 0) };
        session.Execute(new SetCellBordersCommand(addresses, CellBorders.Preset(BorderPreset.All, 1, "FFFF0000")));
        session.Execute(new SetCellBordersCommand(addresses, CellBorders.Preset(BorderPreset.All)));
        Assert.Null(Style(session.Snapshot, 0, 0).TopBorderColor); Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 0));
        Assert.Equal(1, Style(session.Snapshot, 2, 0).TopBorder);
        Assert.Throws<ArgumentException>(() => session.Execute(new SetCellBordersCommand(addresses, new(Top: new(14)))));
    }

    [Fact]
    public void MergedPerimeterAndDiagonalRemainVisibleAndPartialMergeIsAtomic()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Merged", mergedCells: ["B2:D3"]); using var session = new DocumentSession(Book(sheet));
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellBordersCommand([new(sheet.Id, 1, 1)], CellBorders.Preset(BorderPreset.All)))); Assert.Same(before, session.Snapshot);
        session.Execute(new SetCellBordersCommand(Area(sheet.Id, 1, 1, 2, 3), CellBorders.Preset(BorderPreset.All, 6) with { Diagonal = new(9), DiagonalDown = true }));
        var range = new CellRectangle(1, 1, 2, 3);
        Assert.Equal(6, CellBorderProjection.ForRow(range, 2, (r, c) => Style(session.Snapshot, r, c)).RightBorder);
        Assert.Equal(6, CellBorderProjection.ForColumn(range, 3, (r, c) => Style(session.Snapshot, r, c)).BottomBorder);
        Assert.True(CellBorderProjection.ForDiagonals(Style(session.Snapshot, 1, 1)).DiagonalDown);
        Assert.Equal(0, Style(session.Snapshot, 1, 1).RightBorder); Assert.False(Style(session.Snapshot, 2, 3).DiagonalDown);
    }

    [Fact]
    public void FormatDialogCombinesTabsIntoAtomicUndoAndPreservesOtherStyles()
    {
        var sheet = Sheet("S", (1, 1, N(42))); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellFormatCommand(Area(sheet.Id), new(Bold: true, NumberFormatId: 2), CellBorders.Preset(BorderPreset.TopDoubleBottom)));
        Assert.True(Style(session.Snapshot, 2, 2).Bold); Assert.Equal(2, Style(session.Snapshot, 2, 2).NumberFormatId);
        Assert.Equal(6, Style(session.Snapshot, 3, 2).BottomBorder); Assert.Equal(1, Style(session.Snapshot, 1, 2).TopBorder);
        session.Undo(); Assert.False(session.CanUndo); Assert.False(Style(session.Snapshot, 1, 1).Bold); Assert.Equal(0, Style(session.Snapshot, 3, 2).BottomBorder);
    }

    [Fact]
    public void SharedEdgeProjectsTheSameStrongestStrokeOnBothSides()
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 1, 1)], new(RightBorder: 6, RightBorderColor: "FFFF0000")));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 1, 2)], new(LeftBorder: 1, LeftBorderColor: "FF0000FF")));
        var book = session.Snapshot; sheet = book.Sheets[0];
        var left = TablePresentation.GetStyle(book, sheet, 1, 1); var right = TablePresentation.GetStyle(book, sheet, 1, 2);
        Assert.Equal(6, left.RightBorder); Assert.Equal(left.RightBorder, right.LeftBorder); Assert.Equal(left.RightBorderColor, right.LeftBorderColor);
    }

    [Fact]
    public void WholeSheetMergeProjectsOnlyRequestedVisiblePerimeter()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Huge", cells: [new(new(100_000, 0), new(CellContent.FromLiteral(CellValue.Blank, 1)))], mergedCells: [new CellRectangle(0, 0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1).ToString()]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { LeftBorder = 5 }]);
        var reads = 0;
        var range = new CellRectangle(0, 0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1);
        StyleDefinition StyleAt(int r, int c) { reads++; return Style(book, r, c); }
        Assert.Equal(5, CellBorderProjection.ForRow(range, 100_000, StyleAt).LeftBorder);
        Assert.Equal(0, CellBorderProjection.ForRow(range, 100_001, StyleAt).LeftBorder);
        Assert.Equal(0, CellBorderProjection.ForColumn(range, 20, StyleAt).TopBorder);
        Assert.Equal(6, reads);
    }

    [Fact]
    public void MergedRightEdgeDoesNotExtendNeighborBordersAcrossAnUnborderedRow()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Merged", cells: Enumerable.Range(2, 6)
            .Select(r => new KeyValuePair<CellPosition, CellSnapshot>(new(r, 1), new(CellContent.FromLiteral(CellValue.Blank, 1)))), mergedCells: ["A3:A9"]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { LeftBorder = 1 }]);
        var range = CellRectangle.Parse("A3:A9");
        StyleDefinition StyleAt(int r, int c) => TablePresentation.GetStyle(book, sheet, r, c);
        Assert.Equal(1, CellBorderProjection.ForRow(range, 7, StyleAt).RightBorder);
        Assert.Equal(0, CellBorderProjection.ForRow(range, 8, StyleAt).RightBorder);
        Assert.Equal(0, StyleAt(8, 1).LeftBorder);
    }

    [Fact]
    public void RemovingOneDiagonalPreservesTheOtherAndItsStroke()
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellBordersCommand([new(sheet.Id, 0, 0)], new(Diagonal: new(10, "FF0070C0"), DiagonalUp: true, DiagonalDown: true)));
        session.Execute(new SetCellBordersCommand([new(sheet.Id, 0, 0)], new(DiagonalUp: false)));
        var s = Style(session.Snapshot, 0, 0); Assert.False(s.DiagonalUp); Assert.True(s.DiagonalDown); Assert.Equal(10, s.DiagonalBorder); Assert.Equal("FF0070C0", s.DiagonalBorderColor);
    }

    [Fact]
    public void ScriptWeightCannotSilentlyReplaceAnImportedDashDotWithSolid()
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellBordersCommand([new(sheet.Id, 0, 0)], new(Top: new(12)))); var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new SetScriptBorderCommand(sheet.Id, new(0, 0, 0, 0), 0, null, null, 5)));
        Assert.Same(before, session.Snapshot);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void AllLineStylesAndBothDiagonalDirectionsSurviveSaveAndReopen(string extension)
    {
        var sheet = Sheet("S"); using var session = new DocumentSession(Book(sheet));
        for (var kind = 1; kind <= 13; kind++) session.Execute(new SetCellBordersCommand([new(sheet.Id, kind * 2, 1)],
            CellBorders.Preset(BorderPreset.All, kind, "FF123456") with { Diagonal = new(kind, "FF987654"), DiagonalDown = true, DiagonalUp = true }));
        var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(file, session.Snapshot); var opened = storage.Open(file);
            Assert.False(opened.FileState?.FormulaReadOnly == true);
            for (var kind = 1; kind <= 13; kind++)
            {
                var s = Style(opened, kind * 2, 1); Assert.Equal(kind, s.LeftBorder); Assert.Equal(kind, s.RightBorder); Assert.Equal(kind, s.TopBorder); Assert.Equal(kind, s.BottomBorder);
                Assert.Equal(kind, s.DiagonalBorder); Assert.True(s.DiagonalUp && s.DiagonalDown); Assert.Equal("FF987654", s.DiagonalBorderColor); Assert.Equal("FF123456", s.LeftBorderColor);
            }
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void IndependentXlsxWithDashDotAndDiagonalImportsEditable()
    {
        var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            var bytes = WorkbookPackageFixture.Xlsx("<row r='1'><c r='A1' s='0'><v>7</v></c></row>", new Dictionary<string, string> {
                ["xl/styles.xml"] = "<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><fonts><font/></fonts><fills><fill><patternFill patternType='none'/></fill></fills><borders><border diagonalUp='1' diagonalDown='1'><left style='mediumDashDot'><color rgb='FF0070C0'/></left><right/><top/><bottom style='double'/><diagonal style='dashDotDot'><color rgb='FFFF0000'/></diagonal></border></borders><cellXfs><xf fontId='0' fillId='0' borderId='0' numFmtId='0'/></cellXfs></styleSheet>",
                ["xl/_rels/workbook.xml.rels"] = "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet1.xml'/><Relationship Id='styles' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/></Relationships>",
                ["[Content_Types].xml"] = "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/><Override PartName='/xl/worksheets/sheet1.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/><Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/></Types>"
            });
            File.WriteAllBytes(file, bytes); var book = new TypedWorkbookStorage().Open(file); var s = Style(book, 0, 0);
            Assert.Equal(10, s.LeftBorder); Assert.Equal(11, s.DiagonalBorder); Assert.True(s.DiagonalUp && s.DiagonalDown); Assert.False(book.FileState?.FormulaReadOnly == true);
        }
        finally { File.Delete(file); }
    }

    [Theory]
    [InlineData(8)] [InlineData(9)] [InlineData(10)] [InlineData(11)] [InlineData(12)] [InlineData(13)]
    public void SpreadsheetClipboardKeepsNewLinesAndDiagonals(int kind)
    {
        var s = new StyleDefinition { TopBorder = kind, TopBorderColor = "FF123456", DiagonalBorder = kind, DiagonalDown = true, DiagonalUp = true };
        var region = new ExternalClipboardRegion(1, 1, [new(CellValue.FromNumber(1), s)]);
        var result = SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(region));
        Assert.Empty(result.Limitations); Assert.Equal(kind, result.Cells[0].Style!.TopBorder); Assert.True(result.Cells[0].Style!.DiagonalDown && result.Cells[0].Style!.DiagonalUp);
        Assert.NotEmpty(HtmlClipboardCodec.Write(region));
    }
}
