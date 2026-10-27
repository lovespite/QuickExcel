using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class DocumentSessionTests
{
    [Fact]
    public void InvalidFinalAddressRejectsTheWholeTransaction()
    {
        using var session = Session();
        var original = session.Snapshot;
        var notifications = 0;
        session.Changed += (_, _) => notifications++;

        Assert.Throws<ArgumentOutOfRangeException>(() => session.Execute(new SetCellsCommand([
            Edit(session, 0, "valid"), Edit(session, WorkbookLimits.MaxRows, "invalid")])));

        Assert.Same(original, session.Snapshot);
        Assert.Equal(0, session.Revision);
        Assert.False(session.IsDirty);
        Assert.False(session.CanUndo);
        Assert.Equal(0, notifications);
    }

    [Fact]
    public void UnknownSheetAndOversizedFinalTextDoNotLeavePartialEdits()
    {
        using var session = Session();
        var original = session.Snapshot;
        var unknownSheet = new CellEdit(new(SheetId.New(), 0, 0), Text("invalid"));
        Assert.Throws<KeyNotFoundException>(() => session.Execute(new SetCellsCommand([Edit(session, 0, "valid"), unknownSheet])));
        Assert.Throws<ArgumentException>(() => session.Execute(new SetCellsCommand([
            Edit(session, 0, "valid"), Edit(session, 1, new string('x', WorkbookLimits.MaxCellTextLength + 1))])));

        Assert.Same(original, session.Snapshot);
        Assert.Equal(0, session.Revision);
        Assert.False(session.CanUndo);
    }

    [Fact]
    public void CommandAndSheetCapacityAreCheckedBeforeMutation()
    {
        using var session = Session();
        var original = session.Snapshot;
        Assert.Throws<ArgumentException>(() => new SetCellsCommand(
            Enumerable.Repeat(Edit(session, 0, "value"), WorkbookLimits.MaxCells + 1)));
        Assert.Same(original, session.Snapshot);

        using var full = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            Enumerable.Range(0, WorkbookLimits.MaxSheets).Select(index => new WorksheetSnapshot(SheetId.New(), $"Sheet{index}"))));
        var fullSnapshot = full.Snapshot;
        Assert.Throws<ArgumentException>(() => full.Execute(new AddSheetCommand("Overflow")));
        Assert.Same(fullSnapshot, full.Snapshot);
        Assert.False(full.CanUndo);
    }

    [Fact]
    public void DuplicateAddressesAndUnknownStylesRejectTheWholeTransaction()
    {
        using var session = Session();
        var original = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(new SetCellsCommand([Edit(session, 0, "first"), Edit(session, 0, "second")])));
        Assert.Throws<ArgumentException>(() => session.Execute(new SetCellsCommand([
            Edit(session, 0, "first"), new(Address(session, 1), CellContent.FromLiteral(CellValue.FromNumber(1), 1))])));
        Assert.Same(original, session.Snapshot);
    }

    [Fact]
    public void BatchSetAndClearEachUseExactlyOneHistoryEntry()
    {
        using var session = Session();
        var notifications = new List<DocumentChange>();
        session.Changed += (_, change) => notifications.Add(change);
        var edits = Enumerable.Range(0, 20).Select(index => Edit(session, index, $"value-{index}")).ToArray();
        var set = session.Execute(new SetCellsCommand(edits));
        Assert.True(set.Changed);
        Assert.Equal(1, set.Revision);
        Assert.Equal(20, session.Snapshot.Sheets[0].Cells.Count);
        Assert.Equal(20, Assert.Single(notifications).Cells.Length);
        Assert.True(session.Undo());
        Assert.Empty(session.Snapshot.Sheets[0].Cells);
        Assert.False(session.CanUndo);
        Assert.True(session.Redo());
        Assert.Equal(20, session.Snapshot.Sheets[0].Cells.Count);

        session.Execute(new ClearCellsCommand(edits.Select(edit => edit.Address)));
        Assert.Empty(session.Snapshot.Sheets[0].Cells);
        Assert.Equal(4, session.Revision);
        Assert.True(session.Undo());
        Assert.Equal(20, session.Snapshot.Sheets[0].Cells.Count);
        Assert.True(session.Undo());
        Assert.Empty(session.Snapshot.Sheets[0].Cells);
        Assert.False(session.CanUndo);
    }

    [Fact]
    public void ClearPreservesStyleButRemovesUnstyledEmptyCells()
    {
        var sheetId = SheetId.New();
        var cells = new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = new(Text("styled", 1)),
            [new(1, 0)] = new(Text("plain"))
        };
        using var session = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(sheetId, "Sheet1", cells)], styles: [new(), new() { Bold = true }]));
        session.Execute(new ClearCellsCommand([new(sheetId, 0, 0), new(sheetId, 1, 0)]));

        Assert.Single(session.Snapshot.Sheets[0].Cells);
        var styledBlank = session.Snapshot.Sheets[0].GetCell(0, 0)!;
        Assert.Equal(CellValue.Blank, styledBlank.Value);
        Assert.Equal(1, styledBlank.Content.StyleId);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(1, 0));
        Assert.True(session.Undo());
        Assert.Equal(CellValue.FromText("styled"), Cell(session, 0).Value);
        Assert.Equal(CellValue.FromText("plain"), Cell(session, 1).Value);
    }

    [Fact]
    public void NoOpCommandsDoNotConsumeRevisionHistoryOrNotifications()
    {
        using var session = Session();
        session.Execute(new SetCellsCommand([Edit(session, 0, "same")]));
        var before = session.Snapshot;
        var bytes = session.HistoryBytes;
        var notified = false;
        session.Changed += (_, _) => notified = true;

        Assert.False(session.Execute(new SetCellsCommand([Edit(session, 0, "same")])).Changed);
        Assert.False(session.Execute(new ClearCellsCommand([Address(session, 10)])).Changed);
        Assert.False(session.Execute(new UpdateScriptSourceCommand("")).Changed);
        Assert.Same(before, session.Snapshot);
        Assert.Equal(bytes, session.HistoryBytes);
        Assert.False(notified);
    }

    [Fact]
    public void UndoRedoPreserveTypedInputFormulaAndCachedResultWhileRevisionIncreases()
    {
        using var session = Session();
        var expected = new[]
        {
            CellContent.FromLiteral(CellValue.FromText("00123")),
            CellContent.FromLiteral(CellValue.FromNumber(123)),
            CellContent.FromLiteral(CellValue.FromBoolean(true)),
            CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Value)),
            CellContent.FromFormula("=1+2")
        };
        session.Execute(new SetCellsCommand(expected.Select((content, row) => new CellEdit(Address(session, row), content))));
        var request = session.BeginCalculation();
        Assert.True(session.TryCommitCalculation(Result(request, Address(session, 4), CellValue.FromNumber(3))));
        Assert.Equal(1, session.Revision);
        Assert.Equal(CalculationState.Fresh, Cell(session, 4).Calculation!.State);
        Assert.True(session.Undo());
        Assert.Equal(2, session.Revision);
        Assert.Empty(session.Snapshot.Sheets[0].Cells);
        Assert.True(session.Redo());
        Assert.Equal(3, session.Revision);
        for (var row = 0; row < expected.Length; row++) Assert.Equal(expected[row], Cell(session, row).Content);
        Assert.Equal(CellValue.FromNumber(3), Cell(session, 4).Value);
        Assert.Equal(CalculationState.Stale, Cell(session, 4).Calculation!.State);
        Assert.False(session.TryCommitCalculation(Result(request, Address(session, 4), CellValue.FromNumber(99))));
    }

    [Fact]
    public void EditingScriptPreservesCellUndoAndNeverRestoresExecutionPermission()
    {
        using var session = Session();
        session.Execute(new SetCellsCommand([Edit(session, 0, "cell")]));
        session.SetScriptsEnabled(true);
        session.Execute(new UpdateScriptSourceCommand("// new source"));

        Assert.False(session.ScriptsEnabled);
        Assert.True(session.CanUndo);
        Assert.True(session.Undo());
        Assert.Equal("// new source", session.Snapshot.ScriptSource);
        Assert.Empty(session.Snapshot.Sheets[0].Cells);
        Assert.False(session.ScriptsEnabled);
        Assert.True(session.Redo());
        Assert.Equal("// new source", session.Snapshot.ScriptSource);
        Assert.Equal(CellValue.FromText("cell"), Cell(session, 0).Value);
        Assert.False(session.ScriptsEnabled);
    }

    [Fact]
    public void ScriptEditorChangesPreserveCellRedoHistory()
    {
        using var session = Session();
        session.Execute(new SetCellsCommand([Edit(session, 0, "cell")]));
        Assert.True(session.Undo());
        session.Execute(new UpdateScriptSourceCommand("// independent script edit"));

        Assert.True(session.CanRedo);
        Assert.True(session.Redo());
        Assert.Equal("// independent script edit", session.Snapshot.ScriptSource);
        Assert.Equal(CellValue.FromText("cell"), Cell(session, 0).Value);
        Assert.False(session.ScriptsEnabled);
    }

    [Fact]
    public void SavingACapturedRevisionCannotClearLaterDirtyState()
    {
        using var session = Session();
        session.Execute(new SetCellsCommand([Edit(session, 0, "first")]));
        var captured = session.Snapshot;
        session.Execute(new SetCellsCommand([Edit(session, 0, "second")]));
        Assert.True(session.MarkSaved(session.SessionId, captured.Revision));
        Assert.Equal(1, session.SavedRevision);
        Assert.Equal(2, session.Revision);
        Assert.True(session.IsDirty);
        Assert.False(session.MarkSaved(Guid.NewGuid(), session.Revision));
        Assert.False(session.MarkSaved(session.SessionId, session.Revision + 1));
        Assert.False(session.MarkSaved(session.SessionId, 0));
        Assert.True(session.MarkSaved(session.SessionId, session.Revision));
        Assert.False(session.IsDirty);
        Assert.True(session.Undo());
        Assert.True(session.IsDirty);
    }

    [Fact]
    public void SessionsAndCloseReopenRejectResultsFromThePreviousOwner()
    {
        var original = FormulaBook();
        using var first = new DocumentSession(original);
        using var second = new DocumentSession(original);
        var firstRequest = first.BeginCalculation();
        second.BeginCalculation();
        var result = Result(firstRequest, Address(first, 0), CellValue.FromNumber(2));
        Assert.NotEqual(first.SessionId, second.SessionId);
        Assert.False(second.TryCommitCalculation(result));
        Assert.Equal(CellValue.Blank, Cell(second, 0).Value);
        first.Dispose();
        using var reopened = new DocumentSession(original);
        reopened.BeginCalculation();
        Assert.False(first.TryCommitCalculation(result));
        Assert.False(reopened.TryCommitCalculation(result));
        Assert.False(first.MarkSaved(first.SessionId, 0));
        Assert.False(first.CanUndo);
        Assert.Throws<ObjectDisposedException>(() => first.Execute(new UpdateScriptSourceCommand("// closed")));
        Assert.Throws<ObjectDisposedException>(() => first.BeginCalculation());
        Assert.Throws<ObjectDisposedException>(() => first.Snapshot);
    }

    [Fact]
    public void StartingANewerCalculationRejectsOutOfOrderCompletion()
    {
        using var session = new DocumentSession(FormulaBook());
        var oldRequest = session.BeginCalculation();
        var currentRequest = session.BeginCalculation();
        Assert.True(currentRequest.Stamp.CalculationGeneration > oldRequest.Stamp.CalculationGeneration);
        Assert.True(session.TryCommitCalculation(Result(currentRequest, Address(session, 0), CellValue.FromNumber(2))));
        Assert.False(session.TryCommitCalculation(Result(oldRequest, Address(session, 0), CellValue.FromNumber(99))));
        Assert.Equal(CellValue.FromNumber(2), Cell(session, 0).Value);
        Assert.Equal(0, session.Revision);
        Assert.False(session.IsDirty);
    }

    [Fact]
    public void ScriptSourceReapplyDisableAndInputChangesInvalidateEarlierRequests()
    {
        using var session = new DocumentSession(FormulaBook());
        session.SetScriptsEnabled(true);
        var enabled = session.BeginCalculation();
        session.Execute(new UpdateScriptSourceCommand("// updated"));
        Assert.False(session.TryCommitCalculation(Result(enabled, Address(session, 0), CellValue.FromNumber(1))));
        var sourceChanged = session.BeginCalculation();
        Assert.True(sourceChanged.Stamp.ScriptGeneration > enabled.Stamp.ScriptGeneration);
        Assert.False(sourceChanged.ScriptsEnabled);

        session.SetScriptsEnabled(true);
        var applied = session.BeginCalculation();
        session.SetScriptsEnabled(true);
        Assert.False(session.TryCommitCalculation(Result(applied, Address(session, 0), CellValue.FromNumber(1))));
        var reapplied = session.BeginCalculation();
        session.SetScriptsEnabled(false);
        Assert.False(session.TryCommitCalculation(Result(reapplied, Address(session, 0), CellValue.FromNumber(1))));
        var disabled = session.BeginCalculation();
        session.Execute(new SetCellsCommand([Edit(session, 1, "input changed")]));
        Assert.False(session.TryCommitCalculation(Result(disabled, Address(session, 0), CellValue.FromNumber(1))));
    }

    [Fact]
    public void AFailedCurrentCalculationPublishesDiagnosticsWithoutGrantingExecution()
    {
        using var session = new DocumentSession(FormulaBook());
        session.SetScriptsEnabled(true);
        var request = session.BeginCalculation();
        Assert.True(session.TryFailCalculation(request.Stamp, "failure", disableScripts: true));
        Assert.False(session.ScriptsEnabled);
        Assert.Equal("failure", Assert.Single(session.Diagnostics));
        Assert.False(session.TryCommitCalculation(Result(request, Address(session, 0), CellValue.FromNumber(2))));
        Assert.False(session.TryFailCalculation(request.Stamp, "old failure", disableScripts: true));
        Assert.Equal("failure", Assert.Single(session.Diagnostics));
    }

    [Fact]
    public void InvalidCalculationPatchRejectsEveryCellAndKeepsDiagnostics()
    {
        using var session = new DocumentSession(FormulaBook());
        session.Execute(new SetCellsCommand([Edit(session, 1, "literal")]));
        var request = session.BeginCalculation();
        var original = session.Snapshot;
        var patch = new Dictionary<CellAddress, CellCalculation>
        {
            [Address(session, 0)] = Calculation(request, CellValue.FromNumber(2)),
            [Address(session, 1)] = Calculation(request, CellValue.FromNumber(3))
        };
        Assert.Throws<InvalidOperationException>(() => session.TryCommitCalculation(new(request.Stamp, patch, ["not published"])));
        Assert.Same(original, session.Snapshot);
        Assert.Empty(session.Diagnostics);

        patch.Remove(Address(session, 1));
        patch[new(SheetId.New(), 0, 0)] = Calculation(request, CellValue.FromNumber(3));
        Assert.Throws<KeyNotFoundException>(() => session.TryCommitCalculation(new(request.Stamp, patch)));
        Assert.Same(original, session.Snapshot);
        Assert.True(session.TryCommitCalculation(Result(request, Address(session, 0), CellValue.FromNumber(2))));
    }

    [Fact]
    public void PerCellStampMismatchAndExcessiveCachedTextRejectCalculationAtomically()
    {
        using var session = new DocumentSession(FormulaBook());
        var request = session.BeginCalculation();
        var original = session.Snapshot;
        var wrongVersion = new CellCalculation(CellValue.FromNumber(2), CalculationState.Fresh,
            request.Stamp.InputRevision + 1, request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration);
        Assert.Throws<InvalidOperationException>(() => session.TryCommitCalculation(new(request.Stamp,
            new Dictionary<CellAddress, CellCalculation> { [Address(session, 0)] = wrongVersion })));
        var longText = Result(request, Address(session, 0), CellValue.FromText(new('x', WorkbookLimits.MaxCellTextLength + 1)));
        Assert.Throws<InvalidOperationException>(() => session.TryCommitCalculation(longText));
        Assert.Same(original, session.Snapshot);
    }

    [Fact]
    public void AVisibleHistoryBudgetDiscardsOldestEntriesAndRejectsOneOversizedEdit()
    {
        using var session = Session(historyBudgetBytes: 1_800);
        session.Execute(new SetCellsCommand([Edit(session, 0, "1")]));
        session.Execute(new SetCellsCommand([Edit(session, 0, "2")]));
        session.Execute(new SetCellsCommand([Edit(session, 0, "3")]));
        Assert.Equal(2, session.DiscardedHistoryCount);
        Assert.InRange(session.HistoryBytes, 1, session.HistoryBudgetBytes);
        Assert.True(session.Undo());
        Assert.Equal(CellValue.FromText("2"), Cell(session, 0).Value);
        Assert.False(session.CanUndo);
        var before = session.Snapshot;
        var bytes = session.HistoryBytes;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([Edit(session, 0, new string('x', 1_000))])));
        Assert.Same(before, session.Snapshot);
        Assert.Equal(bytes, session.HistoryBytes);
        Assert.True(session.CanRedo);
        Assert.True(session.Redo());
        Assert.Equal(CellValue.FromText("3"), Cell(session, 0).Value);
    }

    [Fact]
    public async Task ParallelWritersReceiveDistinctSerializedRevisions()
    {
        using var session = Session();
        var commands = Enumerable.Range(0, 40).Select(index => new SetCellsCommand([Edit(session, index, index.ToString())])).ToArray();
        var results = await Task.WhenAll(commands.Select(command => Task.Run(() => session.Execute(command))));

        Assert.Equal(Enumerable.Range(1, 40).Select(value => (long)value), results.Select(result => result.Revision).Order());
        Assert.Equal(40, session.Revision);
        Assert.Equal(40, session.Snapshot.Sheets[0].Cells.Count);
        Assert.All(results, result => Assert.True(result.Changed));
    }

    [Fact]
    public async Task ExpectedRevisionAndCancellationCannotWriteStaleCommands()
    {
        using var session = Session();
        var commands = new[] { new SetCellsCommand([Edit(session, 0, "first")]), new SetCellsCommand([Edit(session, 1, "second")]) };
        var accepted = await Task.WhenAll(commands.Select(command => Task.Run(() =>
        {
            try { session.Execute(command, expectedRevision: 0); return true; }
            catch (InvalidOperationException) { return false; }
        })));
        Assert.Single(accepted, value => value);
        Assert.Equal(1, session.Revision);
        Assert.Single(session.Snapshot.Sheets[0].Cells);
        var before = session.Snapshot;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => session.ExecuteAsync(
            new SetCellsCommand([Edit(session, 2, "cancelled")]), cancellationToken: new CancellationToken(true)));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public async Task ConcurrentNotificationsRemainOrderedAndCarryTheirOwnCommittedSnapshot()
    {
        using var session = Session();
        var enteredFirstNotification = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var releaseFirstNotification = new ManualResetEventSlim();
        var notifications = new List<DocumentChange>();
        session.Changed += (_, change) =>
        {
            notifications.Add(change);
            if (change.Revision != 1) return;
            enteredFirstNotification.SetResult();
            if (!releaseFirstNotification.Wait(TimeSpan.FromSeconds(10)))
                throw new TimeoutException("The test did not release the first notification.");
        };
        var first = Task.Run(() => session.Execute(new SetCellsCommand([Edit(session, 0, "first")])));
        try
        {
            await enteredFirstNotification.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await Task.Run(() => session.Execute(new SetCellsCommand([Edit(session, 0, "second")]))).WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(2, session.Revision);
            Assert.Equal(1, Assert.Single(notifications).Revision);
        }
        finally { releaseFirstNotification.Set(); }
        await first.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(new long[] { 1, 2 }, notifications.Select(change => change.Revision));
        Assert.All(notifications, change =>
        {
            Assert.Equal(session.SessionId, change.SessionId);
            Assert.NotNull(change.Snapshot);
            Assert.Equal(change.Revision, change.Snapshot.Revision);
        });
        Assert.Equal(CellValue.FromText("first"), notifications[0].Snapshot!.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText("second"), notifications[1].Snapshot!.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void ReentrantNotificationCommandsAreQueuedAndObserverFailuresDoNotUndoCommits()
    {
        using var session = Session();
        var notifications = new List<long>();
        var depth = 0;
        var maximumDepth = 0;
        session.Changed += (_, _) => throw new InvalidOperationException("A broken observer.");
        session.Changed += (_, change) =>
        {
            depth++;
            maximumDepth = Math.Max(maximumDepth, depth);
            notifications.Add(change.Revision);
            if (change.Revision == 1) session.Execute(new SetCellsCommand([Edit(session, 1, "nested")]));
            depth--;
        };

        var result = session.Execute(new SetCellsCommand([Edit(session, 0, "outer")]));

        Assert.True(result.Changed);
        Assert.Equal(1, result.Revision);
        Assert.Equal(2, session.Revision);
        Assert.Equal(1, maximumDepth);
        Assert.Equal(new long[] { 1, 2 }, notifications);
        Assert.Equal(CellValue.FromText("outer"), Cell(session, 0).Value);
        Assert.Equal(CellValue.FromText("nested"), Cell(session, 1).Value);
    }

    [Fact]
    public void MarkSavedNotifiesWithoutCreatingAnInputRevisionOrHistoryEntry()
    {
        using var session = Session();
        session.Execute(new SetCellsCommand([Edit(session, 0, "saved")]));
        var notifications = new List<DocumentChange>();
        session.Changed += (_, change) => notifications.Add(change);
        var snapshot = session.Snapshot;
        var historyBytes = session.HistoryBytes;

        Assert.True(session.MarkSaved(session.SessionId, session.Revision));

        var notification = Assert.Single(notifications);
        Assert.Equal(1, notification.Revision);
        Assert.Equal(session.SessionId, notification.SessionId);
        Assert.Same(snapshot, notification.Snapshot);
        Assert.Empty(notification.Cells);
        Assert.Same(snapshot, session.Snapshot);
        Assert.Equal(historyBytes, session.HistoryBytes);
        Assert.False(session.IsDirty);
        Assert.True(session.CanUndo);
        Assert.False(session.MarkSaved(Guid.NewGuid(), session.Revision));
        Assert.Single(notifications);
    }

    [Fact]
    public void CompletedCalculationStampCannotCommitOrFailAgain()
    {
        using var session = new DocumentSession(FormulaBook());
        var request = session.BeginCalculation();
        Assert.True(session.TryCommitCalculation(Result(request, Address(session, 0), CellValue.FromNumber(2))));
        var completed = session.Snapshot;

        Assert.False(session.TryCommitCalculation(Result(request, Address(session, 0), CellValue.FromNumber(99))));
        Assert.False(session.TryFailCalculation(request.Stamp, "late failure", disableScripts: true));
        Assert.Same(completed, session.Snapshot);
        Assert.Equal(CellValue.FromNumber(2), Cell(session, 0).Value);
        Assert.Empty(session.Diagnostics);
    }

    [Fact]
    public void CancelledCalculationStampCannotPublishLateSuccessEvenWhenScriptsStayEnabled()
    {
        using var session = new DocumentSession(FormulaBook());
        session.SetScriptsEnabled(true);
        var request = session.BeginCalculation();
        Assert.True(session.TryFailCalculation(request.Stamp, "cancelled", disableScripts: false));
        var cancelled = session.Snapshot;

        Assert.True(session.ScriptsEnabled);
        Assert.False(session.TryCommitCalculation(Result(request, Address(session, 0), CellValue.FromNumber(99))));
        Assert.False(session.TryFailCalculation(request.Stamp, "duplicate cancellation", disableScripts: false));
        Assert.Same(cancelled, session.Snapshot);
        Assert.Equal("cancelled", Assert.Single(session.Diagnostics));
        var retry = session.BeginCalculation();
        Assert.True(session.TryCommitCalculation(Result(retry, Address(session, 0), CellValue.FromNumber(2))));
    }

    private static DocumentSession Session(long historyBudgetBytes = 32 * 1024 * 1024) => new(
        new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new WorksheetSnapshot(SheetId.New(), "Sheet1")]),
        historyBudgetBytes);

    private static WorkbookSnapshot FormulaBook() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new WorksheetSnapshot(SheetId.New(), "Sheet1", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = new(CellContent.FromFormula("=1+1"))
        })]);

    private static CellContent Text(string value, int styleId = 0) => CellContent.FromLiteral(CellValue.FromText(value), styleId);
    private static CellAddress Address(DocumentSession session, int row) => new(session.Snapshot.Sheets[0].Id, row, 0);
    private static CellEdit Edit(DocumentSession session, int row, string value) => new(Address(session, row), Text(value));
    private static CellSnapshot Cell(DocumentSession session, int row) => session.Snapshot.Sheets[0].GetCell(row, 0)!;
    private static CellCalculation Calculation(CalculationRequest request, CellValue value) => new(value,
        value.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh,
        request.Stamp.InputRevision, request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration);
    private static CalculationResult Result(CalculationRequest request, CellAddress address, CellValue value) =>
        new(request.Stamp, new Dictionary<CellAddress, CellCalculation> { [address] = Calculation(request, value) });
}
