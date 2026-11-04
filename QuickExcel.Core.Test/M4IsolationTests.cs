using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Core.Test;

#pragma warning disable CA1416 // Windows x64 integration suite.
public sealed class M4IsolationTests
{
    [Fact]
    public void ActualWorkerHasZeroCapabilitiesAndUniqueIdentityAndCleansOwnedResources()
    {
        var first = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create("host.functions.register('F',()=>42);");
        var directory = first.SandboxPayloadDirectory;
        var profile = first.SandboxProfileName;
        try
        {
            using var process = Process.GetProcessById(first.ProcessId);
            Marshal.ThrowExceptionForHR(DeriveAppContainerSidFromAppContainerName(profile, out var sid));
            try { WorkerProcess.VerifyToken(process.SafeHandle, sid); }
            finally { FreeSid(sid); }
            Assert.Equal(42d, first.Invoke("F", []));
            using var second = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create("host.functions.register('F',()=>7);");
            Assert.NotEqual(profile, second.SandboxProfileName);
            Assert.NotEqual(directory, second.SandboxPayloadDirectory);
            Assert.Equal(7d, second.Invoke("F", []));
        }
        finally { first.Dispose(); }
        Assert.False(Directory.Exists(directory));
        Assert.False(Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages", profile)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProductionLauncherDeniesPrivateFilesPayloadWritesChildProcessesAndUnlistedHandles(bool compilation)
    {
        var folder = Path.Combine(Path.GetTempPath(), "QuickExcel.IsolationTest." + Guid.NewGuid().ToString("N"));
        var acl = new DirectorySecurity();
        acl.SetAccessRuleProtection(true, false);
        acl.AddAccessRule(new(WindowsIdentity.GetCurrent().User!, FileSystemRights.FullControl,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
        new DirectoryInfo(folder).Create(acl);
        var path = Path.Combine(folder, "canary.txt");
        File.WriteAllText(path, "QuickExcel private canary");
        try
        {
            var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Inherit = true };
            using var canary = CreateFile(path, 0x80000000, 3, ref security, 3, 0x80, IntPtr.Zero);
            Assert.False(canary.IsInvalid);
            // The parent can read this inheritable handle; the launcher's three-handle allowlist must exclude it.
            var buffer = new byte[128];
            Assert.True(ReadFile(canary, buffer, buffer.Length, out var count, IntPtr.Zero));
            Assert.True(count > 0);
            Assert.True(SetFilePointerEx(canary, 0, out _, 0));
            var executable = Path.Combine(AppContext.BaseDirectory, "WorkerFixture", "WorkerFixture.exe");
            string json;
            if (compilation)
            {
                // Fault fixture deliberately speaks the test protocol. Exercise
                // the real compile launch policy without adding OS APIs to JSC.
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                using var worker = WorkerProcess.Start(executable, deadline.Token, compilation: true);
                var identity = M4ProtocolTests.Identity();
                WorkerEnvelope Exchange<T>(string kind, T payload)
                {
                    identity = identity with { RequestId = Guid.NewGuid() };
                    WorkerProtocol.WriteAsync(worker.Input, WorkerProtocol.Message(identity, kind, payload), deadline.Token).GetAwaiter().GetResult();
                    var reply = WorkerProtocol.ReadAsync(worker.Output, deadline.Token).GetAwaiter().GetResult();
                    WorkerProtocol.ValidateReply(reply, identity, kind); return reply;
                }
                Exchange("hello", new WorkerHello(WorkerProtocol.Engine, WorkerProtocol.Policy, [3], WorkerProtocol.MaxFrameBytes));
                Exchange("load", new LoadScript("isolation", "fixture"));
                var reply = WorkerProtocol.Payload<BatchResult>(Exchange("invokeBatch", new InvokeBatch([new("F", [WireValue.FromObject(path), WireValue.FromObject(canary.DangerousGetHandle().ToInt64().ToString())])])));
                json = Assert.IsType<string>(Assert.Single(reply.Results).Value!.ToObject());
            }
            else
            {
                using var runtime = new ScriptWorkerRuntimeFactory(executable).Create("isolation");
                json = Assert.IsType<string>(runtime.Invoke("F", [path, canary.DangerousGetHandle().ToInt64().ToString()]));
            }
            var results = JsonSerializer.Deserialize<Dictionary<string, bool>>(json)!;
            Assert.All(results, result => Assert.True(result.Value, result.Key));
            Assert.Equal(8, results.Count);
            Assert.Equal("QuickExcel private canary", File.ReadAllText(path));
        }
        finally { File.Delete(path); Directory.Delete(folder); }
    }

    [Fact]
    public void EnvironmentIsAnAllowlistAndPreCancelledLaunchCreatesNothing()
    {
        using var sandbox = WorkerSandbox.Create(Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe"), default);
        var names = sandbox.EnvironmentBlock().Split('\0', StringSplitOptions.RemoveEmptyEntries).Select(e => e.Split('=')[0]).ToArray();
        Assert.Equal(["DOTNET_CLI_TELEMETRY_OPTOUT", "DOTNET_EnableDiagnostics", "LOCALAPPDATA", "SystemRoot", "TEMP", "TMP", "WINDIR"], names);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => WorkerProcess.Start(sandbox.Executable, cancelled.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationAndForcedExitStillCleanProfileAndPayload(bool forceExit)
    {
        using var cancellation = new CancellationTokenSource();
        var runtime = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(
            "host.functions.register('F',()=>42);", new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), cancellation.Token));
        var directory = runtime.SandboxPayloadDirectory;
        var profile = runtime.SandboxProfileName;
        try
        {
            if (forceExit)
            {
                using var process = Process.GetProcessById(runtime.ProcessId); process.Kill(); process.WaitForExit();
                Assert.Throws<ScriptRuntimeFailureException>(() => runtime.Invoke("F", []));
            }
            else
            {
                cancellation.Cancel();
                Assert.ThrowsAny<OperationCanceledException>(() => runtime.Invoke("F", []));
            }
        }
        finally { runtime.Dispose(); }
        Assert.False(Directory.Exists(directory));
        Assert.False(Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages", profile)));
    }

    [StructLayout(LayoutKind.Sequential)] private struct SecurityAttributes { public int Length; public IntPtr Descriptor; [MarshalAs(UnmanagedType.Bool)] public bool Inherit; }
    [DllImport("userenv.dll", CharSet = CharSet.Unicode)] private static extern int DeriveAppContainerSidFromAppContainerName(string name, out IntPtr sid);
    [DllImport("advapi32.dll")] private static extern IntPtr FreeSid(IntPtr sid);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateFileW")] private static extern SafeFileHandle CreateFile(string path, uint access, uint share, ref SecurityAttributes security, uint creation, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadFile(SafeFileHandle handle, byte[] buffer, int count, out int read, IntPtr overlapped);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetFilePointerEx(SafeFileHandle handle, long distance, out long position, uint method);
}
