using System.Text;
using QuickExcel.Clipboard;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Calculation;

namespace QuickExcel.Core.Test;

public sealed class ExternalClipboardTests
{
    private static byte[] Xml(string cells, int rows = 1, int columns = 1, string styles = "", string extra = "") => Encoding.UTF8.GetBytes(
        $"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Styles>{styles}</Styles><Worksheet ss:Name=\"S\"><Table ss:ExpandedRowCount=\"{rows}\" ss:ExpandedColumnCount=\"{columns}\">{cells}</Table>{extra}</Worksheet></Workbook>\0");
    private static WorkbookSnapshot Book(WorksheetSnapshot? sheet = null, ExcelDateSystem system = ExcelDateSystem.Windows1900) =>
        new(Guid.NewGuid(), 0, system, [sheet ?? new(SheetId.New(), "Target")]);
    private static ExternalClipboardRegion Region(params CellValue[] values) => new(1, values.Length, values.Select(v => new ExternalClipboardCell(v, new() { Bold = true, NumberFormatCode = "0.00%" })));

    [Fact]
    public void TypedValuesAreNeverReparsedAndRoundTripWithOneUndo()
    {
        var values = new[] { CellValue.FromText("00123"), CellValue.FromText("123"), CellValue.FromText("=1+2"), CellValue.FromText("123456789012345678901234"),
            CellValue.FromNumber(.12), CellValue.FromBoolean(true), CellValue.FromError(CellErrorCode.DivisionByZero), CellValue.FromText("中文\r\n行\t\"列\""), CellValue.Blank };
        var source = SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(Region(values)));
        Assert.Equal(values, source.Cells.Select(c => c.Value));
        var book = Book(); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new PasteExternalCellsCommand(source, id, new(3, 2, 3, 2), 0));
        var pasted = session.Snapshot;
        Assert.Equal(2, pasted.Styles.Length); Assert.Equal(values, Enumerable.Range(2, values.Length).Select(c => pasted.Sheets[0].GetCell(3, c)?.Value ?? CellValue.Blank));
        Assert.True(session.Undo()); Assert.Equal(book.Styles, session.Snapshot.Styles); Assert.Empty(session.Snapshot.Sheets[0].Cells); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); Assert.Equal(pasted.Styles, session.Snapshot.Styles);
        var dir = Path.Combine(AppContext.BaseDirectory, "artifacts", "external-clipboard-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        foreach (var ext in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(dir, "typed." + ext); var storage = new TypedWorkbookStorage(); storage.Save(path, pasted); var reopened = storage.Open(path);
            Assert.Equal(values, Enumerable.Range(2, values.Length).Select(c => reopened.Sheets[0].GetCell(3, c)?.Value ?? CellValue.Blank));
            Assert.True(reopened.Styles[reopened.Sheets[0].GetStyleId(3, 2)].Bold);
        }
        session.Execute(new PasteExternalCellsCommand(source, id, new(5, 2, 5, 2), session.Snapshot.Revision));
        Assert.Equal(2, session.Snapshot.Styles.Length);
    }

    [Theory]
    [InlineData(CellPasteMode.All)] [InlineData(CellPasteMode.Formats)]
    public void MergesAndCoveredContentFollowInternalRules(CellPasteMode mode)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Target", [new(new(4, 4), new(CellContent.FromLiteral(CellValue.FromText("keep"))))]);
        var book = Book(sheet); using var session = new DocumentSession(book);
        var region = new ExternalClipboardRegion(2, 2, Enumerable.Repeat(new ExternalClipboardCell(CellValue.Blank, new() { Bold = true }), 4), [new(0, 0, 1, 1)]);
        session.Execute(new PasteExternalCellsCommand(region, sheet.Id, new(3, 3, 3, 3), 0, mode));
        Assert.Contains("D4:E5", session.Snapshot.Sheets[0].MergedCells);
        Assert.Equal(mode == CellPasteMode.Formats ? CellValue.FromText("keep") : CellValue.Blank, session.Snapshot.Sheets[0].GetCell(4, 4)?.Value ?? CellValue.Blank);
        session.Execute(new PasteExternalCellsCommand(Region(CellValue.Blank), sheet.Id, new(3, 3, 4, 4), session.Snapshot.Revision, CellPasteMode.Formats));
        Assert.Empty(session.Snapshot.Sheets[0].MergedCells);
        Assert.True(session.Undo()); Assert.Contains("D4:E5", session.Snapshot.Sheets[0].MergedCells);
    }

    [Theory]
    [InlineData("merge")] [InlineData("hidden")] [InlineData("rich")] [InlineData("protected")] [InlineData("note")] [InlineData("revision")] [InlineData("bounds")]
    public void RejectionsAreAtomicIncludingStyles(string kind)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Target", kind == "rich" ? [new(new(0, 0), new(CellContent.FromRichText(new([new("rich")]))))] : null,
            rows: kind == "hidden" ? [new(0, new(Hidden: true))] : null, mergedCells: kind == "merge" ? ["A1:B2"] : null,
            data: kind == "protected" ? new() { Protection = new() } : null,
            annotations: kind == "note" ? new(notes: [new(new(0, 0), new("author", new([new("note")])) )]) : null);
        var book = Book(sheet); using var session = new DocumentSession(book);
        Assert.ThrowsAny<Exception>(() => session.Execute(new PasteExternalCellsCommand(Region(CellValue.FromNumber(1), CellValue.FromNumber(2)), sheet.Id,
            kind == "bounds" ? new(WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1) : new(0, 0, 0, 0), kind == "revision" ? 1 : 0)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void ValuesKeepTargetStyleAndScalarFillsWithoutTiling()
    {
        var book = Book(); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new PasteExternalCellsCommand(Region(CellValue.FromText("123")), id, new(1, 1, 2, 3), 0, CellPasteMode.Values));
        Assert.Equal(6, session.Snapshot.Sheets[0].Cells.Count); Assert.Single(session.Snapshot.Styles);
        Assert.All(session.Snapshot.Sheets[0].Cells.Values, c => Assert.Equal(CellValue.FromText("123"), c.Value));
        Assert.Throws<ArgumentException>(() => session.Execute(new PasteExternalCellsCommand(Region(CellValue.Blank, CellValue.Blank), id, new(4, 4, 6, 7), session.Snapshot.Revision)));
    }
    [Fact]
    public void ExternalCommandsCannotBypassValidationAndFormulaReadOnly()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Target", data: new() { Validations = [new([new(0, 0, 0, 1)], ValidationKind.Whole, "=1", "=5")] });
        var book = Book(sheet); using var session = new DocumentSession(book, validation: new WorkbookValidation());
        Assert.Throws<DataValidationException>(() => session.Execute(new PasteExternalCellsCommand(Region(CellValue.FromNumber(2), CellValue.FromNumber(9)), sheet.Id, new(0, 0, 0, 0), 0)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900, 45000.5)] [InlineData(ExcelDateSystem.Mac1904, 43538.5)]
    public void ExplicitDatesConvertCalendarSystem(ExcelDateSystem system, double serial)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml("<Row><Cell><Data ss:Type=\"DateTime\">2023-03-15T12:00:00.000</Data></Cell></Row>"));
        var book = Book(system: system); using var session = new DocumentSession(book);
        session.Execute(new PasteExternalCellsCommand(source, book.Sheets[0].Id, new(0, 0, 0, 0), 0));
        Assert.Equal(serial, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }
    [Fact]
    public void FictitiousLeapDayPreservesSameSystemButRejectsCrossSystem()
    {
        var source = SpreadsheetClipboardCodec.Read(Xml("<Row><Cell><Data ss:Type=\"DateTime\">1900-02-29T12:00:00.000</Data></Cell></Row>"));
        Assert.Equal(60.5, source.Cells[0].Value.NumberValue);
        var roundtrip = SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(source)); Assert.Equal(source.Cells[0].Date, roundtrip.Cells[0].Date);
        var book = Book(system: ExcelDateSystem.Mac1904); using var session = new DocumentSession(book);
        Assert.Throws<NotSupportedException>(() => session.Execute(new PasteExternalCellsCommand(source, book.Sheets[0].Id, new(0, 0, 0, 0), 0))); Assert.Same(book, session.Snapshot);
    }

    [Fact]
    public void XmlSparseRowsTrailingBlanksMergesAndInheritedStyles()
    {
        var styles = "<Style ss:ID=\"Default\"><Font ss:FontName=\"Arial\" ss:Bold=\"1\"/></Style><Style ss:ID=\"s1\"><Alignment ss:WrapText=\"1\"/></Style>";
        var source = SpreadsheetClipboardCodec.Read(Xml("<Column ss:StyleID=\"s1\"/><Row><Cell ss:MergeAcross=\"1\"><Data ss:Type=\"String\">a</Data></Cell></Row><Row ss:Index=\"3\"><Cell ss:Index=\"3\"/></Row>", 4, 4, styles));
        Assert.Equal(16, source.Cells.Length); Assert.Single(source.Merges); Assert.True(source.Cells[0].Style!.Bold); Assert.True(source.Cells[0].Style!.WrapText);
        Assert.Equal(CellValue.Blank, source.Cells[^1].Value);
        var empty = SpreadsheetClipboardCodec.Read(Xml("", 3, 4)); Assert.Equal(12, empty.Cells.Length);
    }

    [Theory]
    [InlineData("<Row><Cell ss:Index=\"0\"/></Row>")]
    [InlineData("<Row><Cell ss:MergeAcross=\"9\"/></Row>")]
    [InlineData("<Row><Cell/><Cell ss:Index=\"1\"/></Row>")]
    [InlineData("<Row><Cell><Data ss:Type=\"Number\">NaN</Data></Cell></Row>")]
    [InlineData("<Row><Cell><Data ss:Type=\"Boolean\">TRUE</Data></Cell></Row>")]
    public void MalformedXmlNeverFallsBack(string cells) => Assert.ThrowsAny<Exception>(() => SpreadsheetClipboardCodec.Read(Xml(cells)));

    [Fact]
    public void XmlBudgetsAndEntitiesAreRejected()
    {
        Assert.ThrowsAny<Exception>(() => SpreadsheetClipboardCodec.Read(Xml("", 20001)));
        Assert.ThrowsAny<Exception>(() => SpreadsheetClipboardCodec.Read(new byte[SpreadsheetClipboardCodec.MaxBytes + 1]));
        Assert.ThrowsAny<Exception>(() => SpreadsheetClipboardCodec.Read(Encoding.UTF8.GetBytes("<!DOCTYPE Workbook [<!ENTITY x SYSTEM 'file:///x'>]><Workbook>&x;</Workbook>")));
        Assert.ThrowsAny<Exception>(() => SpreadsheetClipboardCodec.Read(Xml(new string(' ', 1) + string.Concat(Enumerable.Repeat("<a>", 40)) + string.Concat(Enumerable.Repeat("</a>", 40)))));
        Assert.ThrowsAny<Exception>(() => SpreadsheetClipboardCodec.Read(Xml("<Row><Cell ss:StyleID=\"s\"/></Row>", styles: "<Style ss:ID=\"s\" ss:Parent=\"s\"/>")));
    }

    [Fact]
    public void SourceRichTextAndNotesNeedExplicitValuesAndMissingResultsNeverBecomeBlank()
    {
        var rich = SpreadsheetClipboardCodec.Read(Xml("<Row><Cell><Data ss:Type=\"String\"><b>rich</b></Data><Comment>note</Comment></Cell></Row>"));
        Assert.Throws<NotSupportedException>(() => rich.RequireMode(CellPasteMode.All)); rich.RequireMode(CellPasteMode.Values); Assert.Equal("rich", rich.Cells[0].Value.TextValue);
        var missing = SpreadsheetClipboardCodec.Read(Xml("<Row><Cell ss:Formula=\"=1+2\"/></Row>"));
        Assert.Throws<InvalidOperationException>(() => missing.RequireMode(CellPasteMode.Values)); missing.RequireMode(CellPasteMode.Formulas);
    }

    [Theory]
    [InlineData("=R[-1]C+RC[-1]", "=(F5+E6)")]
    [InlineData("=SUM(R[-2]C[-1]:R[-1]C[-1])", "=SUM(E4:E5)")]
    [InlineData("=IF(RC[-1]=0,\"R1C1\",1)", "=IF((E6=0),\"R1C1\",1)")]
    public void OnlyExplicitFormulaModeTranslatesRelativeSyntax(string formula, string expected)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml($"<Row><Cell ss:Formula=\"{System.Net.WebUtility.HtmlEncode(formula)}\"><Data ss:Type=\"Number\">42</Data></Cell></Row>"));
        var book = Book(); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new PasteExternalCellsCommand(source, id, new(5, 5, 5, 5), 0)); Assert.Equal(CellValue.FromNumber(42), session.Snapshot.Sheets[0].GetCell(5, 5)!.Value);
        session.Execute(new PasteExternalCellsCommand(source, id, new(5, 5, 5, 5), session.Snapshot.Revision, CellPasteMode.Formulas));
        Assert.Equal(expected, session.Snapshot.Sheets[0].GetCell(5, 5)!.Content.FormulaText); Assert.Null(session.Snapshot.Sheets[0].GetCell(5, 5)!.Calculation?.Value);
    }
    [Theory]
    [InlineData("=R1C1")] [InlineData("=R1C[-1]")] [InlineData("=Sheet1!RC")] [InlineData("=UNKNOWN(1)")] [InlineData("=SUM(Name)")] [InlineData("=SEQUENCE(2)")]
    public void UnsupportedFormulasOfferOnlyExplicitResults(string formula)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml($"<Row><Cell ss:Formula=\"{formula}\"><Data ss:Type=\"Number\">7</Data></Cell></Row>"));
        Assert.Throws<NotSupportedException>(() => source.RequireMode(CellPasteMode.Formulas)); source.RequireMode(CellPasteMode.Values);
    }

    [Fact]
    public void HtmlUtf8OffsetsContextStylesAndExplicitDisplayText()
    {
        var context = "<html><head><style>.x{font-weight:700;mso-number-format:'0.00%'}</style></head><body>";
        var table = "<table><tr><td class=x>中文😀</td><td rowspan=2>12.00%</td></tr><tr><td>next<br>line</td></tr></table>";
        var bytes = HtmlClipboardCodec.Wrap(context + table + "</body></html>", context.Length, context.Length + table.Length);
        var source = HtmlClipboardCodec.Read(bytes); Assert.Equal(2, source.Rows); Assert.Equal(2, source.Columns); Assert.Single(source.Merges);
        Assert.True(source.Cells[0].Style!.Bold); Assert.Equal("中文😀", source.Cells[0].Value.TextValue); Assert.Equal("next\nline", source.Cells[2].Value.TextValue);
        Assert.Throws<NotSupportedException>(() => source.RequireMode(CellPasteMode.All)); source.RequireMode(CellPasteMode.Values); source.RequireMode(CellPasteMode.Formats);
        bytes[bytes.AsSpan().IndexOf("StartFragment:"u8) + 14] = (byte)'9'; Assert.ThrowsAny<Exception>(() => HtmlClipboardCodec.Read(bytes));
    }
    [Fact]
    public void HtmlDoesNotLoadOrExecuteContentAndRejectsPartialOrHugeMerges()
    {
        var html = "<script>throw 'executed'</script><table><tr><td><a href='file:///x'>value</a></td></tr></table>";
        var source = HtmlClipboardCodec.Read(HtmlClipboardCodec.Wrap(html, 0, html.Length)); Assert.NotEmpty(source.Limitations); source.RequireMode(CellPasteMode.Values);
        html = "<table><tr><td colspan=20001>huge</td></tr></table>";
        Assert.ThrowsAny<Exception>(() => HtmlClipboardCodec.Read(HtmlClipboardCodec.Wrap(html, 0, html.Length)));
        html = "<table><tr><td>a</td></tr><tr><td>b</td></tr></table>";
        Assert.ThrowsAny<Exception>(() => HtmlClipboardCodec.Read(HtmlClipboardCodec.Wrap(html, html.IndexOf("<tr>"), html.IndexOf("<tr>", 8))));
    }
}
