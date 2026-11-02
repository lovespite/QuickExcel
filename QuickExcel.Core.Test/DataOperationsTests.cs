using System.Collections.Immutable;
using FastWpfGrid;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DataOperationsTests
{
    private static WorkbookServices Services => new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts()));
    private static ColumnFilter Values(params CellValue[] values) => new(values.ToImmutableHashSet());

    [Fact]
    public void MultipleKeysMoveWholeRecordsExcludeHeaderPreserveStabilityAndUndo()
    {
        var sheet = Sheet("D", (2, 1, T("部门")), (2, 2, T("金额")), (3, 1, T("B")), (3, 2, N(20)), (3, 3, T("one")),
            (4, 1, T("A")), (4, 2, N(10)), (4, 3, T("two")), (5, 1, T("A")), (5, 2, N(30)), (5, 3, T("three")),
            (6, 1, T("A")), (6, 2, N(30)), (6, 3, T("four")), (3, 0, T("outside")));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        doc.Session.Execute(new SetCellStyleCommand([new(sheet.Id, 5, 3)], new(Bold: true)));
        doc.Session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 5, 1, size: 44, hidden: true));
        var before = doc.Snapshot;
        doc.Session.Execute(new SortRangeCommand(new(sheet.Id, new(2, 1), new(6, 3), [new(1), new(2, true)], true)));
        Assert.Equal(new[] { "three", "four", "two", "one" }, Enumerable.Range(3, 4).Select(r => doc.GetCellInput(0, r, 3)));
        Assert.Equal("部门", doc.GetCellInput(0, 2, 1)); Assert.Equal("outside", doc.GetCellInput(0, 3, 0));
        Assert.True(doc.Snapshot.Styles[doc.Snapshot.Sheets[0].GetStyleId(3, 3)].Bold);
        Assert.True(doc.Snapshot.Sheets[0].Rows[5].Hidden); Assert.Equal(44, doc.Snapshot.Sheets[0].Rows[5].Height);
        Assert.True(doc.Undo()); Assert.Equal(before.Sheets[0].Cells, doc.Snapshot.Sheets[0].Cells);
        Assert.True(doc.Redo()); Assert.Equal("three", doc.GetCellInput(0, 3, 3));
    }

    [Fact]
    public void SecondaryFormulaKeyIsValidatedEvenWhenPrimaryNeverTies()
    {
        var sheet = Sheet("D", (0, 0, N(2)), (1, 0, N(1)), (0, 1, F("=1")), (1, 1, F("=2")));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        var command = new SortRangeCommand(new(sheet.Id, new(0, 0), new(1, 1), [new(0), new(1)]));
        var before = doc.Snapshot;
        Assert.Throws<InvalidOperationException>(() => doc.Session.Execute(command)); Assert.Same(before, doc.Snapshot); Assert.False(doc.Session.CanUndo);
        doc.Calculate(); doc.Session.Execute(command);
        Assert.Equal("=2", doc.GetCellInput(0, 0, 1));
    }

    [Fact]
    public void InvalidSortDescriptorsFailWithoutHistory()
    {
        var sheet = Sheet("D", (0, 0, N(2)), (1, 0, N(1)));
        using var session = new DocumentSession(Book(sheet));
        foreach (var keys in new SortKey[][] { [], [new(1)], [new(0), new(0)], [new(0, basis: SortBasis.FillColor)], [new(0, customOrder: ["a", "A"])], [new(0, basis: (SortBasis)99)] })
            Assert.Throws<ArgumentException>(() => session.Execute(new SortRangeCommand(new(sheet.Id, new(0, 0), new(1, 0), keys))));
        Assert.False(session.CanUndo); Assert.Equal(0, session.Revision);
    }

    [Fact]
    public void CustomOrderAndColorKeysHaveExplicitOrder()
    {
        var sheet = Sheet("D", (0, 0, T("low")), (1, 0, T("high")), (2, 0, T("medium")), (3, 0, T("other")));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SortRangeCommand(new(sheet.Id, new(0, 0), new(3, 0), [new(0, customOrder: ["high", "medium", "low"])])));
        Assert.Equal("high", session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 3, 0)], new(FillColor: "FFFF0000")));
        session.Execute(new SortRangeCommand(new(sheet.Id, new(0, 0), new(3, 0), [new(0, basis: SortBasis.FillColor, preferredColor: "ffff0000")])));
        Assert.Equal("other", session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.TextValue);
    }

    [Fact]
    public void TypedCandidatesDistinguishBlankEmptyStringNumbersTextBooleansErrors()
    {
        var sheet = Sheet("D", (1, 0, T("")), (2, 0, N(1)), (3, 0, T("1")),
            (4, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromBoolean(true)))),
            (5, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Value)))));
        var candidates = DataFilterEngine.Candidates(Book(sheet), new(Guid.NewGuid(), sheet.Id, new(0, 0, 6, 0), false), ImmutableDictionary<int, ColumnFilter>.Empty, 0);
        Assert.Equal(6, candidates.Count); Assert.Equal(2, candidates.Single(c => c.Key.Kind == CellValueKind.Blank).Count);
        Assert.NotEqual(candidates.Single(c => c.Key == CellValue.FromText("1")).Label, candidates.Single(c => c.Key == CellValue.FromNumber(1)).Label);
    }

    [Fact]
    public void OtherColumnCandidatesPreserveCurrentSelectionsWithZeroMatches()
    {
        var sheet = Sheet("D", (0, 0, T("H")), (1, 0, T("A")), (1, 1, N(10)), (2, 0, T("B")), (2, 1, N(20)));
        var filters = ImmutableDictionary<int, ColumnFilter>.Empty.Add(0, Values(CellValue.FromText("B")))
            .Add(1, new(First: new(FilterOperator.Less, "15")));
        var region = new DataRegion(Guid.NewGuid(), sheet.Id, new(0, 0, 2, 1));
        var candidates = DataFilterEngine.Candidates(Book(sheet), region, filters, 0);
        Assert.Equal(1, candidates.Single(c => c.Key == CellValue.FromText("A")).Count);
        var retained = candidates.Single(c => c.Key == CellValue.FromText("B")); Assert.Equal(0, retained.Count); Assert.True(retained.Selected);
        Assert.Equal(2, DataFilterEngine.Evaluate(Book(sheet), region, filters).Count);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 3)]
    public void SameColumnAndOrConditionsUseNumericValues(bool useOr, int visible)
    {
        var sheet = Sheet("D", (1, 0, N(10)), (2, 0, N(20)), (3, 0, N(30)), (4, 0, T("20")));
        var filter = new ColumnFilter(First: new(FilterOperator.GreaterOrEqual, "15"), Second: new(FilterOperator.LessOrEqual, "25"), UseOr: useOr);
        var hidden = DataFilterEngine.Evaluate(Book(sheet), new(Guid.NewGuid(), sheet.Id, new(0, 0, 4, 0)), ImmutableDictionary<int, ColumnFilter>.Empty.Add(0, filter));
        Assert.Equal(visible, 4 - hidden.Count); Assert.Contains(4, hidden); Assert.DoesNotContain(0, hidden);
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900)]
    [InlineData(ExcelDateSystem.Mac1904)]
    public void DateFilterUsesDateSystemAndFormatInsteadOfDisplayStrings(ExcelDateSystem system)
    {
        var serial = ExcelSerialDate.FromDateTime(new DateTime(2026, 10, 29, 15, 0, 0), system).Serial;
        var sheet = Sheet("D", (1, 0, N(serial)), (2, 0, N(serial)), (3, 0, T("2026-10-29")));
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, system, [sheet]));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 1, 0)], new(NumberFormatId: 14)));
        var filter = new ColumnFilter(First: new(FilterOperator.DateOnOrAfter, "2026-10-29"), Second: new(FilterOperator.DateOnOrBefore, "2026-10-29"));
        var hidden = DataFilterEngine.Evaluate(session.Snapshot, new(Guid.NewGuid(), sheet.Id, new(0, 0, 3, 0)), ImmutableDictionary<int, ColumnFilter>.Empty.Add(0, filter));
        Assert.DoesNotContain(1, hidden); Assert.Contains(2, hidden); Assert.Contains(3, hidden);
    }

    [Fact]
    public void FilterReevaluationRejectsStaleCachesAndNeverPublishesOldHiddenRows()
    {
        var sheet = Sheet("D", (1, 0, F("=1")), (2, 0, F("=2")));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        var view = new SheetViewState(); view.SetRegion(new(Guid.NewGuid(), sheet.Id, new(0, 0, 2, 0)));
        Assert.Throws<InvalidOperationException>(() => view.Apply(doc.Snapshot, 0, Values(CellValue.FromNumber(1)))); Assert.Empty(view.Filters);
        doc.Calculate(); view.Apply(doc.Snapshot, 0, Values(CellValue.FromNumber(1))); Assert.True(view.IsFilteredOut(doc.Snapshot, 0, 2));
        doc.SetCell(0, 2, 0, "=1"); view.RefreshFilter(doc.Snapshot);
        Assert.NotNull(view.FilterError); Assert.False(view.IsFilteredOut(doc.Snapshot, 0, 2));
        Assert.Throws<InvalidOperationException>(() => view.RequireVisibleTargets(doc.Snapshot, 0, [2]));
        doc.Calculate(); view.RefreshFilter(doc.Snapshot); Assert.Null(view.FilterError); Assert.False(view.IsFilteredOut(doc.Snapshot, 0, 2));
    }

    [Fact]
    public void SortingFilteredRegionReappliesOnNewSnapshotUndoRedoAndClearingPreservesManualHidden()
    {
        var sheet = Sheet("D", (1, 0, T("B")), (1, 1, N(30)), (2, 0, T("A")), (2, 1, N(10)), (3, 0, T("A")), (3, 1, N(20)));
        using var doc = new WorkbookDocument(Book(sheet), Services);
        doc.Session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 1, 1, hidden: true));
        var view = new SheetViewState(); view.SetRegion(new(Guid.NewGuid(), sheet.Id, new(0, 0, 3, 1)));
        view.Apply(doc.Snapshot, 0, Values(CellValue.FromText("A"))); Assert.True(view.IsFilteredOut(doc.Snapshot, 0, 1));
        var revision = doc.Snapshot.Revision;
        doc.Session.Execute(new SortRangeCommand(new(sheet.Id, new(0, 0), new(3, 1), [new(0), new(1, true)], true)));
        Assert.Equal(revision + 1, doc.Snapshot.Revision); Assert.True(view.IsFilteredOut(doc.Snapshot, 0, 3)); Assert.False(view.IsFilteredOut(doc.Snapshot, 0, 1));
        doc.Undo(); Assert.True(view.IsFilteredOut(doc.Snapshot, 0, 1)); doc.Redo(); Assert.True(view.IsFilteredOut(doc.Snapshot, 0, 3));
        view.ClearConditions(); Assert.False(view.IsFilteredOut(doc.Snapshot, 0, 3)); Assert.True(doc.Snapshot.Sheets[0].Rows[1].Hidden);
    }

    [Fact]
    public void HundredThousandRowsAndCancellationRespectOriginalAddresses()
    {
        var sheet = Sheet("D", (99_999, 1, T("keep")), (2, 0, T("outside")));
        var snapshot = Book(sheet); var region = new DataRegion(Guid.NewGuid(), sheet.Id, new(2, 1, 99_999, 1));
        var filters = ImmutableDictionary<int, ColumnFilter>.Empty.Add(1, Values(CellValue.FromText("keep")));
        var hidden = DataFilterEngine.Evaluate(snapshot, region, filters);
        Assert.Equal(99_996, hidden.Count); Assert.DoesNotContain(2, hidden); Assert.DoesNotContain(99_999, hidden);
        Assert.Throws<OperationCanceledException>(() => DataFilterEngine.Evaluate(snapshot, region, filters, new(true)));
    }

    [Fact]
    public void HugeSelectionIsOneRectangleAndEnumerationIsExplicitlyBounded()
    {
        var selection = new GridRangeSelection(); selection.Set(new GridCellRange(0, 0, 99_999, 255));
        Assert.Single(selection.Ranges); Assert.Equal(25_600_000, selection.Count); Assert.Equal(100_000, selection.RowCount);
        Assert.True(selection.Contains(new(99_999, 255))); Assert.Throws<InvalidOperationException>(() => selection.ToArray());
        selection.Toggle(new GridCellRange(10, 10, 20, 20)); Assert.Equal(25_600_000 - 121, selection.Count);
        Assert.False(selection.Contains(new(15, 15))); Assert.True(selection.Contains(new(21, 15)));
    }

    [Fact]
    public void VisibleIndexPreservesFrozenAndHiddenModelAddressesAtHundredThousandRows()
    {
        var hidden = Enumerable.Range(0, 100_000).Where(r => r % 3 != 0).ToHashSet();
        var frozen = new HashSet<int> { 0, 1, 3, 6 };
        var series = new SeriesSizes { Count = 100_000, DefaultSize = 31 };
        series.SetExtraordinaryIndexes(hidden, frozen);
        var expected = frozen.Except(hidden).Order().Concat(Enumerable.Range(0, 100_000).Except(hidden).Except(frozen)).ToArray();
        Assert.Equal(expected.Length, series.RealCount);
        for (var i = 0; i < expected.Length; i++) { Assert.Equal(expected[i], series.RealToModel(i)); Assert.Equal(i, series.ModelToReal(expected[i])); }
        Assert.Equal(-1, series.ModelToReal(1));
        series.SetExtraordinaryIndexes([], []); Assert.Equal(99_999, series.RealToModel(99_999));
    }

    [Fact]
    public void HiddenRangeEndpointsAndNonPrefixFreezeProjectOnlyTheRequestedModelInterval()
    {
        var series = new SeriesSizes { Count = 100, DefaultSize = 20 };
        series.SetExtraordinaryIndexes([2, 3, 10, 40, 42], [1, 20, 50]);
        for (var first = 0; first < 100; first++)
        for (var last = first; last < 100; last++)
        {
            var actual = series.GetRealRanges(first, last).SelectMany(r => Enumerable.Range(r.First, r.Last - r.First + 1))
                .Select(series.RealToModel).Order().ToArray();
            Assert.Equal(Enumerable.Range(first, last - first + 1).Where(i => i is not (2 or 3 or 10 or 40 or 42)).ToArray(), actual);
        }
        Assert.Empty(series.GetRealRanges(2, 3));
    }

    [Fact]
    public void SparseDimensionHitTestUsesLogicalScrollIndexAtOverrideBoundary()
    {
        var series = new SeriesSizes { Count = 100, DefaultSize = 20 };
        series.Resize(40, 50);
        Assert.Equal(40, series.GetScrollIndexOnPosition(800));
        Assert.Equal(41, series.GetScrollIndexOnPosition(850));
        Assert.Equal(41, series.GetScrollIndexOnPosition(851));
        series.ReplaceSizeOverrides(new Dictionary<int, int>());
        Assert.Equal(20, series.GetSizeByModelIndex(40));
        Assert.Equal(42, series.GetScrollIndexOnPosition(850));
    }

    [Fact]
    public void FilteredViewportStopsAtVisibleRowsIncludingEmptyAndFrozenOnlySheets()
    {
        var series = new SeriesSizes { Count = 2, DefaultSize = 24 };
        series.SetExtraordinaryIndexes([1], []);
        Assert.Equal(1, series.GetVisibleScrollCount(0, 500));
        Assert.Equal(0, series.GetVisibleScrollCount(1, 500));
        Assert.False(series.IsVisible(1, 0, 500));
        series.SetExtraordinaryIndexes([1], [0]);
        Assert.Equal(0, series.GetVisibleScrollCount(0, 500));
        Assert.True(series.IsVisible(0, 0, 500));
        series.SetExtraordinaryIndexes([0, 1], []);
        Assert.Equal(0, series.GetVisibleScrollCount(0, 500));
        Assert.False(series.IsVisible(0, 0, 500));
        series.SetExtraordinaryIndexes([], []);
        Assert.Equal(2, series.GetVisibleScrollCount(0, 500));
    }

    [Fact]
    public void RandomRectangleOperationsEqualReferenceCellSet()
    {
        var random = new Random(402); var selection = new GridRangeSelection(); var expected = new HashSet<FastGridCellAddress>();
        for (var i = 0; i < 300; i++)
        {
            int top = random.Next(20), left = random.Next(10), bottom = random.Next(top, 20), right = random.Next(left, 10);
            var range = new GridCellRange(top, left, bottom, right); var operation = random.Next(3);
            if (operation == 0) selection.Add(range); else if (operation == 1) selection.Remove(range); else selection.Toggle(range);
            for (var row = top; row <= bottom; row++) for (var column = left; column <= right; column++)
            { var cell = new FastGridCellAddress(row, column); if (operation == 0) expected.Add(cell); else if (operation == 1) expected.Remove(cell); else if (!expected.Remove(cell)) expected.Add(cell); }
            Assert.Equal(expected.Count, selection.Count); Assert.True(expected.SetEquals(selection));
            Assert.Equal(expected.Select(c => c.Row).Distinct().Count(), selection.RowCount);
        }
    }
}
