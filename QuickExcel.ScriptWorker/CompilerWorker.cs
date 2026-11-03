using System.Text.Json;
using QuickExcel.ScriptContracts;
using QuickExcel.Scripting;

internal static class CompilerWorker
{
    public static async Task<int> Run()
    {
        if ((!OperatingSystem.IsWindows() && !OperatingSystem.IsLinux()) || !Console.IsInputRedirected || !Console.IsOutputRedirected) return 125;
        try { LinuxWorkerGuard.Enter(); } catch { return 125; }
        using var input = Console.OpenStandardInput();
        using var output = Console.OpenStandardOutput();
        try
        {
            var hello = await WorkerProtocol.ReadAsync(input);
            var accepted = WorkerProtocol.Payload<WorkerHello>(hello);
            var expected = CompilerProtocol.Hello;
            if (hello.Kind != "hello" || accepted.Engine != expected.Engine || accepted.Policy != expected.Policy ||
                accepted.Role != expected.Role || accepted.Compiler != expected.Compiler ||
                !accepted.ApiVersions.SequenceEqual(expected.ApiVersions) || accepted.MaxFrameBytes != expected.MaxFrameBytes)
                throw new ProtocolException("Compiler handshake mismatch.");
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(hello.Identity, "hello", expected));
            var message = await WorkerProtocol.ReadAsync(input);
            if (message.Kind != "compile" || message.Identity.RequestId == hello.Identity.RequestId || !hello.Identity.SameRound(message.Identity))
                throw new ProtocolException("Compiler role or source revision changed.");
            var request = WorkerProtocol.Payload<CompileRequest>(message);
            request.Validate();
            // A fresh, explicit native stack is required by the frozen compiler
            // policy; the execution worker retains its 256 KiB JS stack policy.
            var result = new TaskCompletionSource<CompilerResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
            {
                try
                {
                    var json = JscCompilerSession.Compile(JsonSerializer.Serialize(new { modules = request.Modules, entries = request.Entries }, WorkerProtocol.Json), CompilerProtocol.Identity);
                    result.SetResult(JsonSerializer.Deserialize<CompilerResult>(json, WorkerProtocol.Json) ?? throw new ProtocolException("Missing compiler output."));
                }
                catch (Exception error) { result.SetException(error); }
            }, 8 * 1024 * 1024) { IsBackground = true, Name = "QuickExcel TypeScript" };
            thread.Start();
            var value = await result.Task;
            var reply = new CompiledReply(CompilerProtocol.Identity, request.Sha256, CompilerProtocol.OutputDigest(value), value);
            CompilerProtocol.ValidateReply(reply, request);
            await WorkerProtocol.WriteAsync(output, WorkerProtocol.Message(message.Identity, "compile", reply));
            return 0; // One snapshot per compiler worker; no retained source or VM.
        }
        catch { return 4; }
    }
}
