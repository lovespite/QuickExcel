using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M4MessageBudgetTests
{
    [Fact]
    public async Task ProtocolAcceptsPayloadAboveOldLimitAndKeepsNewByteBoundary()
    {
        Assert.Equal(16 * 1024 * 1024, WorkerProtocol.MaxFrameBytes);
        Assert.Equal(WorkerProtocol.MaxFrameBytes, JscSession.MaxMessageBytes);
        var text = new string('中', 800_000); // Default JSON escaping exceeds 4 MiB.
        var message = WorkerProtocol.Message(M4ProtocolTests.Identity(), "invokeBatch",
            new InvokeBatch([new("F", [WireValue.FromObject(text)])]));
        using var stream = new MemoryStream();
        await WorkerProtocol.WriteAsync(stream, message);
        Assert.InRange(stream.Length, 4 * 1024 * 1024 + 1, WorkerProtocol.MaxFrameBytes);
        stream.Position = 0;
        var received = WorkerProtocol.Payload<InvokeBatch>(await WorkerProtocol.ReadAsync(stream));
        Assert.Equal(text, received.Calls[0].Arguments[0].ToObject());
    }

    [Fact]
    public void NativeBridgeAndRealWorkerAcceptInputsAboveFourMiB()
    {
        const string source = "quickExcel.functions.register('SIZE',s=>s.length); quickExcel.functions.register('LARGE',()=> 'x'.repeat(5*1024*1024));";
        var input = new string('x', 5 * 1024 * 1024);
        using (var native = new JscSession(source))
        {
            Assert.Equal((double)input.Length, native.Invoke("SIZE", [input]));
            Assert.Equal(input, native.Invoke("LARGE", []));
            var error = Assert.Throws<InvalidOperationException>(() => native.Invoke("SIZE", [new string('中', 3_000_000)]));
            Assert.Contains("16 MiB", error.Message);
        }
        using var worker = new ScriptWorkerRuntimeFactory().Create(source);
        Assert.Equal((double)input.Length, worker.Invoke("SIZE", [input]));
    }

    [Fact]
    public void LargeActionSnapshotCrossesOldLimitAndCommitsOnce()
    {
        var sheet = Sheet("S", Enumerable.Range(0, 160).Select(i => (i, 0, T(new string('x', 32767)))).ToArray());
        var modules = M4ActionTests.Actions("""
            quickExcel.actions.register('Copy',({snapshot,patch})=>{
                const id=snapshot.sheets[0].id;
                patch.setValue(id,'B1',snapshot.getRange(id,'A1:A160').reduce((n,r)=>n+r[0].length,0));
            });
            """);
        using var doc = new WorkbookDocument(Book(sheet).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate();
        var revision = doc.Snapshot.Revision;
        Assert.True(doc.RunAction("Copy").Changed);
        Assert.Equal(revision + 1, doc.Snapshot.Revision);
        Assert.Equal(CellValue.FromNumber(160 * 32767), doc.Snapshot.GetSheet(sheet.Id).GetCell(0, 1)!.Value);
        Assert.True(doc.Undo());
        Assert.Null(doc.Snapshot.GetSheet(sheet.Id).GetCell(0, 1));
    }
}
