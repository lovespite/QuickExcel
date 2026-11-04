using System.Text;
using QuickExcel.Clipboard;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class FilteredClipboardTests
{
    private static byte[] Xml(string names = "<NamedRange ss:Name=\"_FilterDatabase\" ss:RefersTo=\"=Source!R1C1:R2C2\" ss:Hidden=\"1\"/>",
        string extra = "", string cellExtra = "", string amount = "<Data ss:Type=\"Number\">0.12</Data>") => Encoding.UTF8.GetBytes($"""
        <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel">
          <Styles><Style ss:ID="Default"><Font ss:Bold="1"/><NumberFormat ss:Format="0.00%"/></Style></Styles>
          <Worksheet ss:Name="Source"><Names>{names}</Names>
            <Table ss:ExpandedRowCount="2" ss:ExpandedColumnCount="2">
              <Row><Cell><Data ss:Type="String">编号</Data></Cell><Cell><Data ss:Type="String">金额</Data></Cell></Row>
              <Row><Cell><Data ss:Type="String">00123</Data>{cellExtra}</Cell><Cell>{amount}</Cell></Row>
            </Table>{extra}
          </Worksheet>
        </Workbook>
        """);

    [Theory]
    [InlineData("_FilterDatabase")]
    [InlineData("_xlnm._FilterDatabase")]
    public void FilterHeadersPasteTypedCellsAndFormatsWithoutImportingFilterRules(string name)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml($"<NamedRange ss:Name=\"{name}\" ss:RefersTo=\"=Source!R1C1:R2C2\" ss:Hidden=\"1\"/>",
            cellExtra: $"<NamedCell ss:Name=\"{name}\"/>"));
        Assert.True(source.HasFilterMetadata); Assert.Empty(source.Limitations);
        var plan = ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All);
        Assert.Equal(CellPasteMode.All, plan.Mode); Assert.Contains("筛选", plan.Notice); Assert.DoesNotContain("只贴值", plan.Notice);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "Target")]);
        using var session = new DocumentSession(book);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(3, 2, 3, 2), 0, plan.Mode));
        var sheet = session.Snapshot.Sheets[0];
        Assert.Equal(CellValue.FromText("00123"), sheet.GetCell(4, 2)!.Value);
        Assert.Equal(CellValue.FromNumber(.12), sheet.GetCell(4, 3)!.Value);
        Assert.True(session.Snapshot.Styles[sheet.GetStyleId(3, 2)].Bold);
        Assert.Null(sheet.Data.Filter); Assert.Empty(session.Snapshot.DefinedNames);
        Assert.True(session.Undo()); Assert.Empty(session.Snapshot.Sheets[0].Cells); Assert.Equal(book.Styles, session.Snapshot.Styles); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); Assert.Equal(CellValue.FromText("00123"), session.Snapshot.Sheets[0].GetCell(4, 2)!.Value);
    }

    [Theory]
    [InlineData("", "<x:AutoFilter x:Range=\"R1C1:R2C2\"/>", "")]
    [InlineData("", "", "<NamedCell ss:Name=\"_FilterDatabase\"/>")]
    public void FilterElementOrReservedCellMarkerIsEnoughToIdentifySource(string names, string extra, string cellExtra)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml(names, extra, cellExtra));
        Assert.True(source.HasFilterMetadata); Assert.Empty(source.Limitations);
    }

    [Theory]
    [InlineData("<Comment>note</Comment>")]
    [InlineData("<DataValidation/>")]
    [InlineData("<ConditionalFormatting/>")]
    [InlineData("<NamedCell ss:Name=\"CustomName\"/>")]
    public void FilterRegionWithOtherLimitationsKeepsSupportedValuesAndFormats(string extra)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml(cellExtra: extra));
        var plan = ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All);
        Assert.Equal(CellPasteMode.All, plan.Mode); Assert.Contains("支持的部分", plan.Notice);
        Assert.All(source.Limitations, limitation => Assert.Contains(limitation, plan.Notice));
        var style = new StyleDefinition { Italic = true, NumberFormatCode = "0.000" };
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "Target")], styles: [style]);
        using var session = new DocumentSession(book);
        session.Execute(new PasteExternalCellsCommand(source, book.ActiveSheetId, new(0, 0, 0, 0), 0, plan.Mode, allowPartial: true));
        Assert.All(session.Snapshot.Sheets[0].Cells, p => Assert.True(session.Snapshot.Styles[p.Value.Content.StyleId].Bold));
        Assert.Equal(CellValue.FromText("00123"), session.Snapshot.Sheets[0].GetCell(1, 0)!.Value);
        Assert.Equal(CellValue.FromNumber(.12), session.Snapshot.Sheets[0].GetCell(1, 1)!.Value);
    }

    [Theory]
    [InlineData(CellPasteMode.Formats)]
    [InlineData(CellPasteMode.Formulas)]
    public void ExplicitFormatsOrFormulasNeverAutomaticallyBecomeValueWrites(CellPasteMode mode)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml(cellExtra: "<Comment>note</Comment>"));
        Assert.Equal(mode, ExternalClipboardPastePolicy.Prepare(source, mode).Mode);
    }

    [Theory]
    [InlineData("<NamedRange ss:Name=\"CustomName\"/>", "")]
    [InlineData("", "<AutoFilter xmlns=\"urn:unknown\"/>")]
    [InlineData("", "<Comment>note</Comment>")]
    public void OtherComplexRegionsAlsoUsePartialPaste(string names, string extra)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml(names, extra));
        Assert.False(source.HasFilterMetadata);
        var plan = ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All);
        Assert.Equal(CellPasteMode.All, plan.Mode); Assert.Contains("支持的部分", plan.Notice);
    }

    [Fact]
    public void OtherNamesAlongsideFilterMetadataDoNotDiscardSupportedFormats()
    {
        var source = SpreadsheetClipboardCodec.Read(Xml("<NamedRange ss:Name=\"_FilterDatabase\"/><NamedRange ss:Name=\"UserName\"/>"));
        Assert.NotEmpty(source.Limitations);
        Assert.Equal(CellPasteMode.All, ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All).Mode);
    }

    [Fact]
    public void MissingResultsAreReportedAndMalformedNumericRepresentationRemainsInvalid()
    {
        var xml = Encoding.UTF8.GetString(Xml(cellExtra: "<Comment>note</Comment>", amount: ""))
            .Replace("<Cell></Cell>", "<Cell ss:Formula=\"=1+2\"/>");
        var source = SpreadsheetClipboardCodec.Read(Encoding.UTF8.GetBytes(xml));
        Assert.Contains("已跳过", ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All).Notice);
        Assert.Throws<FormatException>(() => SpreadsheetClipboardCodec.Read(Xml(amount: "<Data ss:Type=\"Number\">NaN</Data>")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutomaticValuesCannotBypassTargetProtectionOrHiddenRows(bool hidden)
    {
        var source = SpreadsheetClipboardCodec.Read(Xml(cellExtra: "<Comment>note</Comment>"));
        var plan = ExternalClipboardPastePolicy.Prepare(source, CellPasteMode.All);
        var sheet = new WorksheetSnapshot(SheetId.New(), "Target", rows: hidden ? [new(1, new(Hidden: true))] : null,
            data: hidden ? null : new() { Protection = new() });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
        using var session = new DocumentSession(book);
        Assert.ThrowsAny<InvalidOperationException>(() => session.Execute(new PasteExternalCellsCommand(source, sheet.Id, new(0, 0, 0, 0), 0, plan.Mode, allowPartial: true)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }
}
