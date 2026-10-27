using System.Collections.Concurrent;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class WorkbookFacadeTests
{
    [Fact]
    public async Task SavingACapturedWorkbookLeavesEditsDuringIoDirty()
    {
        using var storage = new GatedStorage();
        using var document = WorkbookDocument.CreateBlank(new(storage, new UnexpectedCalculator()));
        document.SetCell(0, 0, 0, "captured");
        var save = Task.Run(() => document.Save("workbook.xlsq"));
        try
        {
            await storage.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            document.SetCell(0, 0, 0, "edited during save");
            Assert.Equal(2, document.Session.Revision);
        }
        finally { storage.Release.Set(); }
        await save.WaitAsync(TimeSpan.FromSeconds(10));

        var saved = Assert.Single(storage.Saved);
        Assert.Equal(1, saved.Snapshot.Revision);
        Assert.Equal(CellValue.FromText("captured"), saved.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal("edited during save", document.GetCellValue(0, 0, 0));
        Assert.Equal(1, document.Session.SavedRevision);
        Assert.True(document.IsDirty);
    }

    [Fact]
    public async Task QueuedSavesSerializeAndCaptureTheLatestStateAfterAcquiringTheWriter()
    {
        using var storage = new GatedStorage();
        using var document = WorkbookDocument.CreateBlank(new(storage, new UnexpectedCalculator()));
        document.SetCell(0, 0, 0, "first");
        var first = Task.Run(() => document.Save("workbook.xlsq"));
        Task? second = null;
        try
        {
            await storage.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            document.SetCell(0, 0, 0, "second");
            var secondAttempted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            second = Task.Run(() =>
            {
                secondAttempted.SetResult();
                document.Save("workbook.xlsq");
            });
            await secondAttempted.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Single(storage.Saved);
        }
        finally { storage.Release.Set(); }
        await Task.WhenAll(first, second!).WaitAsync(TimeSpan.FromSeconds(10));

        var saves = storage.Saved.ToArray();
        Assert.Equal(new long[] { 1, 2 }, saves.Select(saved => saved.Snapshot.Revision));
        Assert.Equal(CellValue.FromText("first"), saves[0].Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText("second"), saves[1].Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(1, storage.MaximumConcurrentSaves);
        Assert.Equal(2, document.Session.SavedRevision);
        Assert.False(document.IsDirty);
    }

    [Fact]
    public void FailedSaveDoesNotAdvanceSavedRevisionOrEraseDirtyState()
    {
        var storage = new RecordingStorage { SaveFailure = new IOException("Simulated storage failure.") };
        using var document = WorkbookDocument.CreateBlank(new(storage, new UnexpectedCalculator()));
        document.SetCell(0, 0, 0, "unsaved");
        var snapshot = document.Snapshot;

        var exception = Assert.Throws<IOException>(() => document.Save("workbook.xlsq"));

        Assert.Same(storage.SaveFailure, exception);
        Assert.Same(snapshot, document.Snapshot);
        Assert.Equal(0, document.Session.SavedRevision);
        Assert.True(document.IsDirty);
        Assert.True(document.CanUndo);
    }

    [Theory]
    [InlineData("input")]
    [InlineData("source")]
    [InlineData("disable")]
    [InlineData("close")]
    public async Task DelayedCalculationCannotPublishAfterItsDocumentStateChanges(string change)
    {
        using var calculator = new GatedCalculator();
        using var document = FormulaDocument(calculator);
        document.Session.SetScriptsEnabled(true);
        var calculation = Task.Run(() => document.Calculate(CancellationToken.None));
        WorkbookSnapshot? afterChange = null;
        try
        {
            await calculator.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            switch (change)
            {
                case "input": document.SetCell(0, 1, 0, "new input"); break;
                case "source": document.ScriptSource = "// new source"; break;
                case "disable": document.DisableScripts(); break;
                case "close": document.Dispose(); break;
            }
            if (!document.Session.IsClosed) afterChange = document.Snapshot;
        }
        finally { calculator.Release.Set(); }

        Assert.False(await calculation.WaitAsync(TimeSpan.FromSeconds(10)));
        if (change == "close") Assert.True(document.Session.IsClosed);
        else
        {
            Assert.Same(afterChange, document.Snapshot);
            Assert.Equal("", document.GetCellValue(0, 0, 0));
            Assert.Empty(document.ScriptErrors);
        }
    }

    [Fact]
    public async Task CancellationAfterCalculatorReturnsCannotPublishItsResult()
    {
        using var calculator = new GatedCalculator();
        using var document = FormulaDocument(calculator);
        using var cancellation = new CancellationTokenSource();
        document.Session.SetScriptsEnabled(true);
        var calculation = Task.Run(() => document.Calculate(cancellation.Token));
        try
        {
            await calculator.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            cancellation.Cancel();
        }
        finally { calculator.Release.Set(); }

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => calculation);
        Assert.Equal("", document.GetCellValue(0, 0, 0));
        Assert.True(document.ScriptsEnabled);
        Assert.Single(document.ScriptErrors);
        Assert.False(document.Session.TryCommitCalculation(calculator.Result!));
    }

    [Fact]
    public async Task FailureFromAnObsoleteCalculationCannotDisableNewAuthorization()
    {
        using var calculator = new GatedCalculator { Failure = new InvalidOperationException("old compile failure") };
        using var document = FormulaDocument(calculator);
        document.Session.SetScriptsEnabled(true);
        var calculation = Task.Run(() => document.Calculate(CancellationToken.None));
        try
        {
            await calculator.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            document.ScriptSource = "// replacement source";
            document.Session.SetScriptsEnabled(true);
        }
        finally { calculator.Release.Set(); }

        await Assert.ThrowsAsync<InvalidOperationException>(() => calculation);
        Assert.True(document.ScriptsEnabled);
        Assert.Empty(document.ScriptErrors);
        Assert.Equal("// replacement source", document.ScriptSource);
    }

    [Fact]
    public void NormalDocumentOperationsUseOnlyInjectedPortsWithoutStartingCalculation()
    {
        var storage = new RecordingStorage();
        var calculator = new UnexpectedCalculator();
        var services = new WorkbookServices(storage, calculator);
        using var document = WorkbookDocument.CreateBlank(services);
        document.SetCells(0, [(0, 0, "'00123"), (1, 0, "=1+1")]);
        Assert.Equal("00123", document.GetCellInput(0, 0, 0));
        Assert.Equal("=1+1", document.GetCellInput(0, 1, 0));
        Assert.True(document.Undo());
        Assert.True(document.Redo());
        document.ScriptSource = "throw new Error('must not run while editing');";
        document.AddSheet("Second");
        document.Save("state.xlsq");
        document.ExportXlsx("results.xlsx");
        storage.OpenSnapshot = document.Snapshot;
        using var reopened = WorkbookDocument.Open("state.xlsq", services);
        using var demo = WorkbookDocument.CreateDemo(services);

        Assert.Equal(0, calculator.CallCount);
        Assert.Single(storage.Saved);
        Assert.Single(storage.Exported);
        Assert.Same(document.Snapshot, reopened.Snapshot);
        Assert.False(reopened.ScriptsEnabled);
        Assert.False(demo.ScriptsEnabled);
        Assert.False(document.IsDirty);
    }

    private static WorkbookDocument FormulaDocument(IWorkbookCalculator calculator)
    {
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(SheetId.New(), "Sheet1", new Dictionary<CellPosition, CellSnapshot>
            {
                [new(0, 0)] = new(CellContent.FromFormula("=1+1"))
            })]);
        return new(snapshot, new(new RecordingStorage(), calculator));
    }

    private sealed class UnexpectedCalculator : IWorkbookCalculator
    {
        public int CallCount { get; private set; }
        public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default)
        {
            CallCount++;
            throw new InvalidOperationException("Ordinary document operations must not start calculation.");
        }
    }

    private class RecordingStorage : IWorkbookStorage
    {
        public ConcurrentQueue<(string Path, WorkbookSnapshot Snapshot)> Saved { get; } = new();
        public ConcurrentQueue<(string Path, WorkbookSnapshot Snapshot)> Exported { get; } = new();
        public Exception? SaveFailure { get; init; }
        public WorkbookSnapshot? OpenSnapshot { get; set; }
        public WorkbookSnapshot Open(string path) => OpenSnapshot ?? throw new IOException("No test workbook has been configured.");
        public virtual void Save(string path, WorkbookSnapshot snapshot)
        {
            if (SaveFailure is not null) throw SaveFailure;
            Saved.Enqueue((path, snapshot));
        }
        public void Export(string path, WorkbookSnapshot snapshot) => Exported.Enqueue((path, snapshot));
    }

    private sealed class GatedStorage : RecordingStorage, IDisposable
    {
        private int _activeSaves;
        private int _saveCount;
        public TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public ManualResetEventSlim Release { get; } = new();
        public int MaximumConcurrentSaves { get; private set; }
        public override void Save(string path, WorkbookSnapshot snapshot)
        {
            var active = Interlocked.Increment(ref _activeSaves);
            MaximumConcurrentSaves = Math.Max(MaximumConcurrentSaves, active);
            try
            {
                base.Save(path, snapshot);
                if (Interlocked.Increment(ref _saveCount) != 1) return;
                Entered.SetResult();
                if (!Release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("The test did not release storage.");
            }
            finally { Interlocked.Decrement(ref _activeSaves); }
        }
        public void Dispose() => Release.Dispose();
    }

    private sealed class GatedCalculator : IWorkbookCalculator, IDisposable
    {
        public TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public ManualResetEventSlim Release { get; } = new();
        public Exception? Failure { get; init; }
        public CalculationResult? Result { get; private set; }
        public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default)
        {
            Entered.SetResult();
            if (!Release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("The test did not release calculation.");
            if (Failure is not null) throw Failure;
            // Deliberately ignore cancellation here: the facade must still prevent publication.
            return Result = new(request.Stamp, request.Snapshot.Sheets.SelectMany(sheet => sheet.Cells
                .Where(pair => pair.Value.Content.Kind == CellContentKind.Formula)
                .Select(pair => new KeyValuePair<CellAddress, CellCalculation>(new(sheet.Id, pair.Key.Row, pair.Key.Column),
                    new(CellValue.FromNumber(2), CalculationState.Fresh, request.Stamp.InputRevision,
                        request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration)))));
        }
        public void Dispose() => Release.Dispose();
    }
}
