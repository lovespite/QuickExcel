using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace QuickExcel.Runtime;

/// <summary>The only production launch path: zero-capability AppContainer, explicit IPC handles, suspended until contained.</summary>
internal sealed class WorkerProcess : IWorkerProcess
{
    private WorkerSandbox? _sandbox;
    private WorkerJob? _job;
    private Process? _process;
    private SafeProcessHandle? _nativeProcess;
    public Stream Input { get; private set; } = Stream.Null;
    public Stream Output { get; private set; } = Stream.Null;
    public Stream Error { get; private set; } = Stream.Null;
    public int Id => _process!.Id;
    public string ProfileName => _sandbox!.ProfileName;
    public string PayloadDirectory => _sandbox!.PayloadDirectory;

    internal static WorkerProcess Start(string executable, CancellationToken cancellation, string? ownershipRoot = null, bool compilation = false)
    {
        var child = new WorkerProcess();
        var allocated = new List<IntPtr>();
        var attributes = IntPtr.Zero;
        var initialized = false;
        try
        {
            child._sandbox = WorkerSandbox.Create(executable, cancellation, ownershipRoot);
            child._job = new WorkerJob(child._sandbox.JobName, compilation);
            var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Inherit = true };
            Check(CreatePipe(out var stdinRead, out var stdinWrite, ref security, 0));
            using (stdinRead) using (stdinWrite)
            {
                Check(CreatePipe(out var stdoutRead, out var stdoutWrite, ref security, 0));
                using (stdoutRead) using (stdoutWrite)
                {
                    Check(CreatePipe(out var stderrRead, out var stderrWrite, ref security, 0));
                    using (stderrRead) using (stderrWrite)
                    {
                        Check(SetHandleInformation(stdinWrite, 1, 0));
                        Check(SetHandleInformation(stdoutRead, 1, 0));
                        Check(SetHandleInformation(stderrRead, 1, 0));
                        nuint size = 0;
                        _ = InitializeProcThreadAttributeList(IntPtr.Zero, 3, 0, ref size);
                        attributes = Marshal.AllocHGlobal((nint)size);
                        Check(InitializeProcThreadAttributeList(attributes, 3, 0, ref size));
                        initialized = true;
                        void Attribute<T>(nuint key, T value) where T : struct
                        {
                            var pointer = Marshal.AllocHGlobal(Marshal.SizeOf<T>()); allocated.Add(pointer);
                            Marshal.StructureToPtr(value, pointer, false);
                            Check(UpdateProcThreadAttribute(attributes, 0, key, pointer, (nuint)Marshal.SizeOf<T>(), IntPtr.Zero, IntPtr.Zero));
                        }
                        Attribute(0x20002, new Handles { Input = stdinRead.DangerousGetHandle(), Output = stdoutWrite.DangerousGetHandle(), Error = stderrWrite.DangerousGetHandle() });
                        Attribute(0x20009, new SecurityCapabilities { Sid = child._sandbox.Sid });
                        // Assign at creation, closing the host-crash gap between CreateProcess and AssignProcessToJobObject.
                        Attribute(0x2000D, child._job.Handle);
                        var startup = new StartupInfoEx
                        {
                            Info = new StartupInfo { Size = Marshal.SizeOf<StartupInfoEx>(), Flags = 0x100,
                                Input = stdinRead.DangerousGetHandle(), Output = stdoutWrite.DangerousGetHandle(), Error = stderrWrite.DangerousGetHandle() },
                            Attributes = attributes
                        };
                        var environment = Marshal.StringToHGlobalUni(child._sandbox.EnvironmentBlock()); allocated.Add(environment);
                        cancellation.ThrowIfCancellationRequested();
                        Check(CreateProcess(child._sandbox.Executable, new StringBuilder('"' + child._sandbox.Executable + '"' + (compilation ? " --compile" : "")),
                            IntPtr.Zero, IntPtr.Zero, true, 0x80000 | 0x4 | 0x8000000 | 0x400, environment,
                            child._sandbox.PayloadDirectory, ref startup, out var info));
                        child._nativeProcess = new SafeProcessHandle(info.Process, ownsHandle: true);
                        using var thread = new SafeFileHandle(info.Thread, ownsHandle: true);
                        child._process = Process.GetProcessById(info.ProcessId);
                        // Verify the OS token ourselves, before any child code (including the CLR) runs.
                        VerifyToken(child._nativeProcess, child._sandbox.Sid);
                        child._job.VerifyMembership(child._nativeProcess);
                        child.Input = DuplicateStream(stdinWrite, FileAccess.Write);
                        child.Output = DuplicateStream(stdoutRead, FileAccess.Read);
                        child.Error = DuplicateStream(stderrRead, FileAccess.Read);
                        cancellation.ThrowIfCancellationRequested();
                        if (ResumeThread(thread) == uint.MaxValue) throw new Win32Exception(Marshal.GetLastWin32Error());
                    }
                }
            }
            return child;
        }
        catch { child.Dispose(); throw; }
        finally
        {
            if (attributes != IntPtr.Zero) { if (initialized) DeleteProcThreadAttributeList(attributes); Marshal.FreeHGlobal(attributes); }
            foreach (var pointer in allocated) Marshal.FreeHGlobal(pointer);
        }
    }

    private static Stream DuplicateStream(SafeFileHandle handle, FileAccess access)
    {
        var current = GetCurrentProcess();
        Check(DuplicateHandle(current, handle, current, out var duplicate, 0, false, 2));
        try { return new FileStream(duplicate, access, 4096, isAsync: false); }
        catch { duplicate.Dispose(); throw; }
    }

    internal static void VerifyToken(SafeProcessHandle process, IntPtr expectedSid)
    {
        Check(OpenProcessToken(process, 8, out var token));
        using (token)
        {
            IntPtr Information(int kind)
            {
                _ = GetTokenInformation(token, kind, IntPtr.Zero, 0, out var size);
                var buffer = Marshal.AllocHGlobal(size);
                try { Check(GetTokenInformation(token, kind, buffer, size, out _)); return buffer; }
                catch { Marshal.FreeHGlobal(buffer); throw; }
            }
            var flag = Information(29); var sid = IntPtr.Zero; var capabilities = IntPtr.Zero;
            try
            {
                sid = Information(31); capabilities = Information(30);
                if (Marshal.ReadInt32(flag) != 1 || !EqualSid(Marshal.ReadIntPtr(sid), expectedSid) || Marshal.ReadInt32(capabilities) != 0)
                    throw new UnauthorizedAccessException("Worker token is not the requested zero-capability AppContainer.");
            }
            finally { Marshal.FreeHGlobal(flag); Marshal.FreeHGlobal(sid); Marshal.FreeHGlobal(capabilities); }
        }
    }

    public void Kill()
    {
        _job?.Dispose();
        if (_nativeProcess is { IsClosed: false, IsInvalid: false }) _ = TerminateProcess(_nativeProcess, 1);
    }

    public void Dispose()
    {
        Kill();
        // Never remove an identity or its payload while its process could still be running.
        if (_nativeProcess is { IsClosed: false, IsInvalid: false } && WaitForSingleObject(_nativeProcess, 1000) != 0)
            throw new IOException("Worker did not exit within its cleanup deadline.");
        Input.Dispose(); Output.Dispose(); Error.Dispose();
        _process?.Dispose(); _nativeProcess?.Dispose();
        _sandbox?.Dispose();
    }

    private static void Check(bool success) { if (!success) throw new Win32Exception(Marshal.GetLastWin32Error()); }
    [StructLayout(LayoutKind.Sequential)] private struct SecurityAttributes { public int Length; public IntPtr Descriptor; [MarshalAs(UnmanagedType.Bool)] public bool Inherit; }
    [StructLayout(LayoutKind.Sequential)] private struct Handles { public IntPtr Input, Output, Error; }
    [StructLayout(LayoutKind.Sequential)] private struct SecurityCapabilities { public IntPtr Sid, Capabilities; public uint Count, Reserved; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct StartupInfo
    {
        public int Size; public string? Reserved, Desktop, Title;
        public uint X, Y, XSize, YSize, XChars, YChars, Fill, Flags;
        public ushort Show, ReservedSize; public IntPtr ReservedBytes, Input, Output, Error;
    }
    [StructLayout(LayoutKind.Sequential)] private struct StartupInfoEx { public StartupInfo Info; public IntPtr Attributes; }
    [StructLayout(LayoutKind.Sequential)] private struct ProcessInformation { public IntPtr Process, Thread; public int ProcessId, ThreadId; }
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool CreatePipe(out SafeFileHandle read, out SafeFileHandle write, ref SecurityAttributes security, uint size);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetHandleInformation(SafeFileHandle handle, uint mask, uint flags);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool InitializeProcThreadAttributeList(IntPtr list, int count, uint flags, ref nuint size);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool UpdateProcThreadAttribute(IntPtr list, uint flags, nuint key, IntPtr value, nuint size, IntPtr previous, IntPtr returned);
    [DllImport("kernel32.dll")] private static extern void DeleteProcThreadAttributeList(IntPtr list);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateProcessW")] private static extern bool CreateProcess(string application, StringBuilder command, IntPtr processAttributes, IntPtr threadAttributes, bool inherit, uint flags, IntPtr environment, string directory, ref StartupInfoEx startup, out ProcessInformation info);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint ResumeThread(SafeFileHandle thread);
    [DllImport("kernel32.dll")] private static extern bool TerminateProcess(SafeProcessHandle process, uint exit);
    [DllImport("kernel32.dll")] private static extern uint WaitForSingleObject(SafeProcessHandle process, uint milliseconds);
    [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool DuplicateHandle(IntPtr sourceProcess, SafeFileHandle source, IntPtr targetProcess, out SafeFileHandle target, uint access, bool inherit, uint options);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool OpenProcessToken(SafeProcessHandle process, uint access, out SafeAccessTokenHandle token);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool GetTokenInformation(SafeAccessTokenHandle token, int kind, IntPtr value, int size, out int needed);
    [DllImport("advapi32.dll")] private static extern bool EqualSid(IntPtr first, IntPtr second);
}
