using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class QuickConditionalFormatTests
{
    private static readonly DifferentialStyle Red = new(FillColor: "FFFFC7CE", FontColor: "FF9C0006");
    private static bool Highlighted(WorkbookSnapshot book, int row, int column = 0) =>
        ConditionalPresentation.GetStyle(book, book.Sheets[0], row, column).FillColor == Red.FillColor;

    [Theory]
    [InlineData(ConditionalPreset.TopItems, 2, "2,3,4")]
    [InlineData(ConditionalPreset.BottomItems, 2, "0,1")]
    [InlineData(ConditionalPreset.TopPercent, 40, "2,3,4")]
    [InlineData(ConditionalPreset.BottomPercent, 21, "0,1")]
    [InlineData(ConditionalPreset.AboveAverage, 10, "2,3,4")]
    [InlineData(ConditionalPreset.BelowAverage, 10, "0,1")]
    [InlineData(ConditionalPreset.TopItems, 10, "0,1,2,3,4")]
    [InlineData(ConditionalPreset.BottomPercent, 100, "0,1,2,3,4")]
    public void RankingUsesNumbersAndIncludesTies(ConditionalPreset preset, int rank, string expected)
    {
        var sheet = Sheet("Data", (0, 0, N(1)), (1, 0, N(2)), (2, 0, N(8)), (3, 0, N(8)), (4, 0, N(10)),
            (5, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("99")))), (6, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromBoolean(true)))));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [ConditionalFormatPresets.Create(new(0, 0, 7, 0), preset, Red, rank: rank)]));
        Assert.Equal(expected, string.Join(',', Enumerable.Range(0, 8).Where(row => Highlighted(session.Snapshot, row))));
    }

    [Fact]
    public void RankingReevaluatesChangedDependenciesAndRoundTripsAsPortableFormula()
    {
        var sheet = Sheet("Data", (0, 0, N(2)), (1, 0, F("=C1*2")), (0, 2, N(3)));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [ConditionalFormatPresets.Create(new(0, 0, 1, 0), ConditionalPreset.TopItems, Red, rank: 1)]));
        Assert.False(Highlighted(session.Snapshot, 0)); Assert.True(Highlighted(session.Snapshot, 1));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 2), N(0).Content)]));
        Assert.True(Highlighted(session.Snapshot, 0)); Assert.False(Highlighted(session.Snapshot, 1));
        session.Undo(); Assert.True(Highlighted(session.Snapshot, 1));
        foreach (var extension in new[] { ".xlsx", ".xlsq" })
        {
            var path = Path.Combine(Path.GetTempPath(), "QuickExcel-quick-cf-" + Guid.NewGuid().ToString("N") + extension);
            try
            {
                var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
                Assert.Empty(reopened.FileState!.CompatibilityDiagnostics);
                Assert.True(session.Snapshot.Sheets[0].Data.ConditionalFormats[0].ContentEquals(reopened.Sheets[0].Data.ConditionalFormats[0]));
                Assert.True(Highlighted(reopened, 1));
            }
            finally { File.Delete(path); }
        }
    }

    [Fact]
    public void QuickComparisonAndTextPresetsHandleValuesFormulasAndQuotes()
    {
        var sheet = Sheet("Data", (0, 0, N(3)), (1, 0, N(7)), (2, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("a\"b")))));
        using var session = new DocumentSession(Book(sheet));
        void Apply(ConditionalPreset preset, string first, string second = "0") => session.Execute(new SetConditionalFormatsCommand(sheet.Id,
            [ConditionalFormatPresets.Create(new(0, 0, 2, 0), preset, Red, first, second)]));
        Apply(ConditionalPreset.Between, "2", "4"); Assert.True(Highlighted(session.Snapshot, 0)); Assert.False(Highlighted(session.Snapshot, 1));
        Apply(ConditionalPreset.GreaterThan, "=$A$1"); Assert.False(Highlighted(session.Snapshot, 0)); Assert.True(Highlighted(session.Snapshot, 1));
        Apply(ConditionalPreset.LessThan, "5"); Assert.True(Highlighted(session.Snapshot, 0)); Assert.False(Highlighted(session.Snapshot, 1));
        Apply(ConditionalPreset.Equal, "a\"b"); Assert.True(Highlighted(session.Snapshot, 2));
        Apply(ConditionalPreset.ContainsText, "\"b"); Assert.False(Highlighted(session.Snapshot, 0)); Assert.True(Highlighted(session.Snapshot, 2));
        Assert.Throws<ArgumentException>(() => Apply(ConditionalPreset.ContainsText, ""));
        Assert.Throws<ArgumentException>(() => Apply(ConditionalPreset.GreaterThan, " "));
        Assert.Throws<ArgumentException>(() => ConditionalFormatPresets.Create(new(0, 0, 1, 0), ConditionalPreset.TopPercent, Red, rank: 101));
    }

    [Fact]
    public void EmptyRankingIsUnformattedWithoutDiagnostics()
    {
        var sheet = Sheet("Data");
        var rule = ConditionalFormatPresets.Create(new(0, 0, 3, 0), ConditionalPreset.TopItems, Red);
        var book = Book(sheet.WithData(new() { ConditionalFormats = [rule] }));
        Assert.False(Highlighted(book, 0)); Assert.Empty(ConditionalPresentation.Diagnostics(book));
    }

    [Fact]
    public void PartialClearSplitsAreasRebasesBothFormulasAndKeepsIdentityAndHistory()
    {
        var sheet = Sheet("Data", (0, 0, N(1)), (1, 0, N(2)), (2, 0, N(3)), (3, 0, N(4)), (4, 0, N(5)));
        var rule = new ConditionalFormatRule([new(0, 0, 4, 2)], ConditionalFormatKind.CellIs, 3, Red, "=A1-1", "=$A1+1");
        using var session = new DocumentSession(Book(sheet.WithData(new() { ConditionalFormats = [rule] })));
        session.Execute(new ClearConditionalFormatsCommand(sheet.Id, [new(0, 0, 0, 2), new(2, 1, 3, 1)]));
        var remaining = session.Snapshot.Sheets[0].Data.ConditionalFormats.Single();
        Assert.Equal(rule.ScriptId, remaining.ScriptId); Assert.Equal(3, remaining.Priority);
        Assert.Equal("=(A2-1)", remaining.Formula1); Assert.Equal("=($A2+1)", remaining.Formula2);
        Assert.Equal(4, remaining.Ranges.Length); Assert.Equal(10L, remaining.Ranges.Sum(r => r.Count));
        Assert.False(remaining.Contains(0, 0)); Assert.False(remaining.Contains(2, 1));
        Assert.True(Highlighted(session.Snapshot, 1)); Assert.True(Highlighted(session.Snapshot, 4));
        session.Undo(); Assert.True(rule.ContentEquals(session.Snapshot.Sheets[0].Data.ConditionalFormats.Single()));
        session.Redo(); Assert.False(Highlighted(session.Snapshot, 0));
        var unchanged = session.Snapshot;
        session.Execute(new ClearConditionalFormatsCommand(sheet.Id, [new(20, 20, 21, 21)])); Assert.Same(unchanged, session.Snapshot);
        session.Execute(new ClearConditionalFormatsCommand(sheet.Id)); Assert.Empty(session.Snapshot.Sheets[0].Data.ConditionalFormats);
        session.Undo(); Assert.Equal(remaining.ScriptId, session.Snapshot.Sheets[0].Data.ConditionalFormats.Single().ScriptId);
    }

    [Fact]
    public void ClearHandlesOverlappingSelectionsAndRemovalOfFirstArea()
    {
        var sheet = Sheet("Data");
        var rule = new ConditionalFormatRule([new(0, 0, 1, 0), new(5, 0, 8, 0)], ConditionalFormatKind.Expression, 1, Red, "=A1>0");
        using var session = new DocumentSession(Book(sheet.WithData(new() { ConditionalFormats = [rule] })));
        session.Execute(new ClearConditionalFormatsCommand(sheet.Id, [new(0, 0, 5, 0), new(1, 0, 6, 0)]));
        var remaining = session.Snapshot.Sheets[0].Data.ConditionalFormats.Single();
        Assert.Equal(new(7, 0, 8, 0), Assert.Single(remaining.Ranges)); Assert.Equal("=(A8>0)", remaining.Formula1);
        session.Execute(new SetProtectionCommand(sheet.Id, new())); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ClearConditionalFormatsCommand(sheet.Id))); Assert.Same(before, session.Snapshot);
    }
}
