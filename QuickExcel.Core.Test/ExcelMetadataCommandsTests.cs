using System.Collections.Immutable;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ExcelMetadataCommandsTests
{
    private static readonly IgnoredCellErrors TextNumber = IgnoredCellErrors.NumberStoredAsText;
    private static WorkbookSnapshot Book(params WorksheetSnapshot[] sheets) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, sheets);
    private static WorksheetSnapshot Sheet(string name, CellRectangle range) => new(SheetId.New(), name,
        [new(new(range.Top, range.Left), new(CellContent.FromLiteral(CellValue.FromText("00123"))))],
        rows: [new(range.Top, new RowMetadata { Descent = .3 })], data: new()
        {
            DefaultRowDescent = .25,
            IgnoredErrors = [new(range, TextNumber)],
            PhoneticProperties = new("<font xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><name val=\"Calibri\"/><sz val=\"11\"/></font>", "noConversion")
        });
    private static IgnoredCellErrors At(WorksheetSnapshot sheet, int row, int col) => sheet.Data.IgnoredErrors
        .Where(e => e.Range.Contains(row, col)).Aggregate(IgnoredCellErrors.None, (flags, e) => flags | e.Errors);

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(-.1)]
    public void InvalidDescentRejectedBeforeSnapshotCanBePublished(double descent)
    {
        Assert.Throws<ArgumentException>(() => new WorksheetSnapshot(SheetId.New(), "S", rows: [new(0, new RowMetadata { Descent = descent })]));
        Assert.Throws<ArgumentException>(() => new WorksheetSnapshot(SheetId.New(), "S", data: new() { DefaultRowDescent = descent }));
    }

    [Fact]
    public void MetadataParticipatesInEqualityAndRangeValidation()
    {
        var data = Sheet("S", new(1, 1, 2, 2)).Data;
        Assert.False(data.IsEmpty);
        Assert.False(data.ContentEquals(data with { DefaultRowDescent = .3 }));
        Assert.False(data.ContentEquals(data with { IgnoredErrors = [] }));
        Assert.False(data.ContentEquals(data with { PhoneticProperties = null }));
        Assert.Throws<ArgumentException>(() => new WorksheetSnapshot(SheetId.New(), "S", data: new() { IgnoredErrors = [new(new(2, 0, 1, 0), TextNumber)] }));
        Assert.Throws<ArgumentException>(() => new WorksheetSnapshot(SheetId.New(), "S", data: new() { IgnoredErrors = [new(new(0, 0, 0, 0), (IgnoredCellErrors)512)] }));
        _ = new WorksheetSnapshot(SheetId.New(), "S", data: new() { IgnoredErrors = [new(new(0, 0, 0, 0), IgnoredCellErrors.None)] });
    }

    [Fact]
    public void PhoneticFontIdentityUsesExpandedNamesAndNormalizedXml()
    {
        var a = new WorksheetPhoneticProperties("<font xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><color tint='0.5' theme='1'/></font>");
        var b = new WorksheetPhoneticProperties("<x:font xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>\n<x:color theme='1' tint='0.5' />\n</x:font>");
        Assert.Equal(a, b);
        Assert.Equal(a, b with { FontXml = a.FontXml });
        Assert.NotEqual(a, new WorksheetPhoneticProperties("<font xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><color theme='2' tint='0.5'/></font>"));
        Assert.Throws<ArgumentException>(() => new WorksheetPhoneticProperties("<!DOCTYPE font [<!ENTITY x SYSTEM 'file:///missing'>]><font>&x;</font>"));
    }

    [Fact]
    public void InsertDeleteAndUndoMapRangesAndRowDescentTogether()
    {
        var sheet = Sheet("S", new(1, 1, 3, 2));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 2));
        var current = session.Snapshot.GetSheet(sheet.Id);
        Assert.Equal(new CellRectangle(3, 1, 5, 2), Assert.Single(current.Data.IgnoredErrors).Range);
        Assert.Equal(.3, current.Rows[3].Descent);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Column, 1, 1, delete: true));
        Assert.Equal(new CellRectangle(3, 1, 5, 1), Assert.Single(session.Snapshot.GetSheet(sheet.Id).Data.IgnoredErrors).Range);
        Assert.True(session.Undo()); Assert.True(session.Undo());
        Assert.Equal(sheet.Data, session.Snapshot.GetSheet(sheet.Id).Data);
        Assert.Equal(.3, session.Snapshot.GetSheet(sheet.Id).Rows[1].Descent);
    }

    [Fact]
    public void CrossSheetCutSplitsSourceAndReplacesDestinationErrorPreferencesAtomically()
    {
        var source = Sheet("Source", new(0, 0, 2, 2));
        var destination = Sheet("Destination", new(4, 4, 6, 6));
        destination = destination.WithData(destination.Data with { IgnoredErrors = [new(new(4, 4, 6, 6), IgnoredCellErrors.Formula)] });
        using var session = new DocumentSession(Book(source, destination));
        session.Execute(new MoveCellsCommand(source.Id, new(1, 1), new(1, 1), destination.Id, new(5, 5)));
        var movedSource = session.Snapshot.GetSheet(source.Id); var movedTarget = session.Snapshot.GetSheet(destination.Id);
        Assert.Equal(IgnoredCellErrors.None, At(movedSource, 1, 1));
        Assert.Equal(TextNumber, At(movedSource, 1, 0));
        Assert.Equal(TextNumber, At(movedTarget, 5, 5));
        Assert.Equal(IgnoredCellErrors.Formula, At(movedTarget, 4, 4));
        Assert.True(session.Undo());
        Assert.Equal(source.Data, session.Snapshot.GetSheet(source.Id).Data);
        Assert.Equal(destination.Data, session.Snapshot.GetSheet(destination.Id).Data);
    }

    [Fact]
    public void TransposedAndScalarFillPasteMapPreferencesButValuesPasteRetainsTarget()
    {
        var source = Sheet("S", new(0, 0, 1, 0));
        using var session = new DocumentSession(Book(source));
        var clipboard = WorkbookClipboard.Capture(session.Snapshot, source.Id, new(0, 0, 1, 1));
        session.Execute(new PasteCellsCommand(clipboard, source.Id, new(3, 3, 4, 4), session.Snapshot.Revision, transpose: true));
        Assert.Equal(TextNumber, At(session.Snapshot.GetSheet(source.Id), 3, 4));
        Assert.Equal(IgnoredCellErrors.None, At(session.Snapshot.GetSheet(source.Id), 4, 3));
        clipboard = WorkbookClipboard.Capture(session.Snapshot, source.Id, new(0, 0, 0, 0));
        session.Execute(new PasteCellsCommand(clipboard, source.Id, new(6, 6, 7, 7), session.Snapshot.Revision));
        Assert.Equal(TextNumber, At(session.Snapshot.GetSheet(source.Id), 7, 7));
        clipboard = WorkbookClipboard.Capture(session.Snapshot, source.Id, new(8, 8, 8, 8));
        session.Execute(new PasteCellsCommand(clipboard, source.Id, new(7, 7, 7, 7), session.Snapshot.Revision, CellPasteMode.Values));
        Assert.Equal(TextNumber, At(session.Snapshot.GetSheet(source.Id), 7, 7));
    }

    [Fact]
    public void CopySheetRetainsFontPayloadAndLayoutHints()
    {
        var source = Sheet("S", new(0, 0, 2, 2));
        using var session = new DocumentSession(Book(source));
        var command = new CopySheetCommand(source.Id, "Copy", 1); session.Execute(command);
        var copy = session.Snapshot.GetSheet(command.Id);
        Assert.True(source.Data.ContentEquals(copy.Data));
        Assert.Equal(source.Rows[0], copy.Rows[0]);
    }

    [Fact]
    public void PartialPreferenceSortFailsWithoutChangingSnapshot()
    {
        var source = Sheet("S", new(1, 0, 1, 0));
        using var session = new DocumentSession(Book(source)); var before = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new SortRangeCommand(source.Id, new(0, 0), new(2, 0), 0)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void XlsqMetadataRoundTripsAndLegacyPackageGainsFeatureWhenSaved()
    {
        var folder = Path.Combine(Path.GetTempPath(), "QuickExcel-excel-metadata-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var file = Path.Combine(folder, "metadata.xlsq"); var source = Sheet("S", new(1, 1, 3, 2));
            var storage = new TypedWorkbookStorage(); storage.Save(file, Book(source));
            var reopened = storage.Open(file).Sheets[0];
            Assert.True(source.Data.ContentEquals(reopened.Data)); Assert.Equal(.3, reopened.Rows[1].Descent);
            using (var archive = ZipFile.Open(file, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("manifest.json")!; JsonObject manifest;
                using (var reader = new StreamReader(entry.Open())) manifest = JsonNode.Parse(reader.ReadToEnd())!.AsObject();
                var features = manifest["requiredFeatures"]!.AsArray();
                Assert.Contains(features, n => n!.GetValue<string>() == "worksheet-excel-metadata");
                var feature = features.Single(n => n!.GetValue<string>() == "worksheet-excel-metadata"); features.Remove(feature); entry.Delete();
                using var writer = new StreamWriter(archive.CreateEntry("manifest.json").Open(), new UTF8Encoding(false)); writer.Write(manifest.ToJsonString());
            }
            // Older writers retained these XML properties inside their original workbook.xlsx
            // without knowing the new feature. They remain valid v2 input.
            var legacy = storage.Open(file);
            Assert.True(source.Data.ContentEquals(legacy.Sheets[0].Data));
            storage.Save(file, legacy);
            using var verified = ZipFile.OpenRead(file);
            using var finalReader = new StreamReader(verified.GetEntry("manifest.json")!.Open());
            var finalManifest = JsonNode.Parse(finalReader.ReadToEnd())!;
            Assert.Contains(finalManifest["requiredFeatures"]!.AsArray(), n => n!.GetValue<string>() == "worksheet-excel-metadata");
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public void SubtotalsRetainInertLayoutAndPhoneticMetadataButRejectIgnoredErrorRanges()
    {
        var original = Sheet("S", new(0, 0, 2, 1));
        CellSnapshot Text(string value) => new(CellContent.FromLiteral(CellValue.FromText(value)));
        CellSnapshot Number(double value) => new(CellContent.FromLiteral(CellValue.FromNumber(value)));
        var cells = new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = Text("Group"), [new(0, 1)] = Text("Value"),
            [new(1, 0)] = Text("a"), [new(1, 1)] = Number(1),
            [new(2, 0)] = Text("a"), [new(2, 1)] = Number(2)
        };
        original = original.WithCells(cells.ToImmutableDictionary());
        Assert.Throws<NotSupportedException>(() => SubtotalCommands.Preview(Book(original), original.Id, new(0, 0, 2, 1), 0, [1]));
        var source = original.WithData(original.Data with { IgnoredErrors = [] });
        using var session = new DocumentSession(Book(source));
        var preview = SubtotalCommands.Preview(session.Snapshot, source.Id, new(0, 0, 2, 1), 0, [1]);
        Assert.Equal(2, preview.InsertedRows);
        session.Execute(new ApplySubtotalsCommand(preview));
        var result = session.Snapshot.GetSheet(source.Id);
        Assert.Equal(source.Data.DefaultRowDescent, result.Data.DefaultRowDescent);
        Assert.Equal(source.Data.PhoneticProperties, result.Data.PhoneticProperties);
        Assert.Equal(source.Rows[0].Descent, result.Rows[0].Descent);
        Assert.True(session.Undo()); Assert.True(source.Data.ContentEquals(session.Snapshot.GetSheet(source.Id).Data));
    }
}
