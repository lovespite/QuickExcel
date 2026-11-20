using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptConsoleTests
{
    private static WorkbookDocument Document() => new(Book(Sheet("Data")).WithScriptSource("throw Error('UNTRUSTED WORKBOOK MUST NOT RUN');"),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));

    [Fact]
    public void ImmediateExpressionsPrintObjectsAndNeverExecuteOrAuthorizeStoredModules()
    {
        using var doc = Document(); var source = doc.ScriptSource; var revision = doc.Snapshot.Revision;
        Assert.False(doc.RunConsole("console.log('中文', {answer:42}); console.info('info'); console.warn('warn'); console.error('error'); console.debug('debug'); 2 + 3").Changed);
        var output = doc.Session.ConsoleEntries.Where(e => e.Kind == "output").Select(e => e.Text).ToArray();
        Assert.Contains("中文 {answer: 42}", output); Assert.Contains("=> 5", output);
        foreach (var level in new[] { "info", "warn", "error", "debug" }) Assert.Contains($"[{level}] {level}", output);
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Level == ScriptConsoleLevel.Warning && e.DisplayText == "warn");
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Level == ScriptConsoleLevel.Error && e.DisplayText == "error");
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Level == ScriptConsoleLevel.Info && e.DisplayText == "info");
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Level == ScriptConsoleLevel.Debug && e.DisplayText == "debug");
        Assert.False(doc.ScriptsEnabled); Assert.Equal(source, doc.ScriptSource); Assert.Equal(revision, doc.Snapshot.Revision); Assert.False(doc.IsDirty);
    }

    [Fact]
    public void ConsolePatchesAndVirtualFilesCommitTogetherAndUndoWithoutChangingSource()
    {
        using var doc = Document(); var source = doc.ScriptSource;
        Assert.True(doc.RunConsole("workbook.getActiveWorksheet().getCell('A1').setValue(42); host.files.writeText('//xlsq/console.txt','ok'); host.files.writeText('//mem/session','one');").Changed);
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.True(doc.Snapshot.PackageFiles.Entries.ContainsKey("//xlsq/console.txt"));
        Assert.True(doc.Undo()); Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.Empty(doc.Snapshot.PackageFiles.Entries);
        Assert.True(doc.Redo()); Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Equal(source, doc.ScriptSource);
        doc.RunConsole("host.files.readText('//mem/session')"); Assert.Contains(doc.Session.ConsoleEntries, e => e.Text == "=> one");
        Assert.False(doc.ScriptsEnabled);
    }

    [Fact]
    public void ExceptionKeepsPriorLogsButRollsBackEveryPatch()
    {
        using var doc = Document();
        var error = Assert.Throws<ScriptRuntimeFailureException>(() => doc.RunConsole("console.log('before failure'); workbook.getActiveWorksheet().getCell('A1').setValue(42); host.files.writeText('//mem/lost','x'); throw Error('STOP');"));
        Assert.Contains("before failure", error.Logs);
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Text == "before failure");
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Kind == "output" && e.Level == ScriptConsoleLevel.Error && e.Text.Contains("STOP"));
        Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.Empty(doc.Session.TransientFiles.Entries); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void TimeoutDoesNotCommitPartialWritesAndNextCommandGetsAFreshWorker()
    {
        using var doc = Document();
        Assert.Throws<TimeoutException>(() => doc.RunConsole("workbook.getActiveWorksheet().getCell('A1').setValue(42); while (true) {}"));
        Assert.Equal("", doc.GetCellValue(0, 0, 0)); Assert.False(doc.CanUndo);
        doc.RunConsole("7 * 6"); Assert.Contains(doc.Session.ConsoleEntries, e => e.Text == "=> 42");
    }

    [Fact]
    public void ConsoleHasNoHostGlobalsAndVariablesArePerCommand()
    {
        using var doc = Document();
        doc.RunConsole("const privateVariable=5; [typeof document,typeof process,typeof require,typeof fetch]");
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Text == "=> [undefined, undefined, undefined, function]");
        Assert.ThrowsAny<Exception>(() => doc.RunConsole("fetch('https://example.com/')"));
        Assert.False(doc.ScriptsEnabled);
        doc.RunConsole("typeof privateVariable"); Assert.Contains(doc.Session.ConsoleEntries, e => e.Text == "=> undefined");
        doc.RunConsole("""const circular={value:1}; circular.self=circular; console.log(circular); 'quotes " and slash \\';""");
        Assert.Contains(doc.Session.ConsoleEntries, e => e.Text.Contains("[Circular]"));
    }

    [Fact]
    public void StaleConsoleTransactionsCannotBypassRevisionOrCancellationGuards()
    {
        using var doc = Document(); var request = doc.Session.BeginConsole("1");
        doc.SetCell(0, 0, 0, "newer");
        var result = new ScriptActionResult(request.OperationId, request.Stamp.InputRevision, request.Name!,
            [new(new(doc.Snapshot.ActiveSheetId, 0, 0), CellValue.FromNumber(42), null)], []);
        Assert.Throws<OperationCanceledException>(() => doc.Session.CommitAction(request, result));
        Assert.Equal("newer", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void ConsoleCannotRewriteScriptsOutsideUndoableDataHistory()
    {
        using var doc = Document(); var source = doc.ScriptSource;
        Assert.Throws<InvalidOperationException>(() => doc.RunConsole("host.files.writeText('//xlsq/scripts/main.js', '// overwritten');"));
        Assert.Equal(source, doc.ScriptSource); Assert.False(doc.CanUndo);
    }

    [Fact]
    public void HostRejectionRemainsVisibleAfterTheLogBudgetIsFull()
    {
        using var doc = Document();
        Assert.Throws<InvalidOperationException>(() => doc.RunConsole("for(let i=0;i<100;i++) console.log('line'); host.files.writeText('//xlsq/scripts/main.js','// changed');"));
        Assert.Contains("控制台不能改写", doc.Session.ConsoleEntries[^1].Text);
        Assert.Equal(ScriptConsoleLevel.Error, doc.Session.ConsoleEntries[^1].Level);
        Assert.False(doc.CanUndo);
    }

    [Theory]
    [InlineData("failure-log-count")]
    [InlineData("failure-log-bytes")]
    public void MalformedWorkerFailureLogsAreRejectedBeforePresentation(string fault)
    {
        var snapshot = Book(Sheet("Data")).WithScriptModules(M4ActionTests.Actions(fault));
        var executor = new ScriptWorkerActionExecutor(Path.Combine(AppContext.BaseDirectory, "Fixtures", "WorkerFixture", "WorkerFixture.exe"));
        var error = Assert.Throws<ScriptRuntimeFailureException>(() => executor.Run(
            new(new(Guid.NewGuid(), snapshot.Revision, 1, 1), Guid.NewGuid(), "A", snapshot, default), default));
        Assert.Equal("Protocol", error.Code); Assert.Empty(error.Logs);
    }

    [Fact]
    public void ReadOnlyConsoleAllowsInspectionButRejectsCellAndFilePatches()
    {
        var original = Book(Sheet("Data"));
        var state = new WorkbookFileState("artifacts/readonly-console.xlsx", 2, [], original, [], [],
            preservedFormulas: [new(PreservedFormulaKind.Array, "xl/worksheets/sheet1.xml", "Data", "A1", "A1:A2", "SUM(A1)")]);
        using var session = new DocumentSession(original.WithFileState(state));
        var inspect = session.BeginConsole("1+1");
        Assert.False(session.CommitAction(inspect, new(inspect.OperationId, inspect.Stamp.InputRevision, inspect.Name!, [], ["=> 2"])).Changed);
        var write = session.BeginConsole("workbook.getActiveWorksheet().getCell('A1').setValue(42)");
        Assert.Throws<InvalidOperationException>(() => session.CommitAction(write, new(write.OperationId, write.Stamp.InputRevision, write.Name!,
            [new(new(original.ActiveSheetId, 0, 0), CellValue.FromNumber(42), null)], [])));
        var file = session.BeginConsole("host.files.writeText('//mem/data','x')");
        Assert.Throws<InvalidOperationException>(() => session.CommitAction(file, new(file.OperationId, file.Stamp.InputRevision, file.Name!, [], [])
            { FileEdits = [new("write", "//mem/data", System.Collections.Immutable.ImmutableArray.Create<byte>(1))] }));
        Assert.False(session.IsDirty); Assert.False(session.CanUndo); Assert.Empty(session.TransientFiles.Entries);
    }

    [Fact]
    public void OutputIsBoundedAndClearDoesNotChangeTheWorkbook()
    {
        using var doc = Document();
        for (var round = 0; round < 6; round++) doc.RunConsole("for(let i=0;i<110;i++) console.log('x'.repeat(5000));");
        Assert.Equal(DocumentSession.MaxConsoleEntries, doc.Session.ConsoleEntries.Length);
        Assert.All(doc.Session.ConsoleEntries, e => Assert.InRange(e.Text.Length, 0, 4000));
        doc.Session.ClearConsole(); Assert.Empty(doc.Session.ConsoleEntries); Assert.False(doc.IsDirty); Assert.False(doc.CanUndo);
        Assert.Throws<ArgumentException>(() => doc.RunConsole(new string('中', DocumentSession.MaxConsoleSourceBytes)));
    }

    [Theory]
    [InlineData("output", "[warn] line\nnext", ScriptConsoleLevel.Warning, "line\nnext")]
    [InlineData("output", "[error] failure", ScriptConsoleLevel.Error, "failure")]
    [InlineData("output", "JavaScript：[warn] formula warning", ScriptConsoleLevel.Warning, "JavaScript：formula warning")]
    [InlineData("output", "JavaScript：[error] formula error", ScriptConsoleLevel.Error, "JavaScript：formula error")]
    [InlineData("output", "a [error] label", ScriptConsoleLevel.Log, "a [error] label")]
    [InlineData("input", "[warn] typed input", ScriptConsoleLevel.Input, "[warn] typed input")]
    [InlineData("output", "=> 42", ScriptConsoleLevel.Result, "42")]
    public void LogLevelsPreserveRawTextAndCalculationContext(string kind, string text, ScriptConsoleLevel level, string display)
    {
        var entry = new ScriptConsoleEntry(1, DateTimeOffset.UtcNow, kind, text);
        Assert.Equal(level, entry.Level); Assert.Equal(display, entry.DisplayText); Assert.Equal(text, entry.Text);
    }
}
