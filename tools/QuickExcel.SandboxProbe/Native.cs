using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using QuickExcel.Runtime;

namespace QuickExcel.SandboxProbe;

internal static class Native
{
    private static void Check(bool success) { if (!success) throw new Win32Exception(Marshal.GetLastWin32Error()); }
    internal static void Rewind(SafeFileHandle handle) => Check(SetFilePointerEx(handle, 0, out _, 0));

    internal static SafeFileHandle OpenInheritableFile(string path)
    {
        var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Inherit = true };
        var handle = CreateFile(path, 0x80000000, 3, ref security, 3, 0x80, IntPtr.Zero);
        if (handle.IsInvalid) { handle.Dispose(); throw new Win32Exception(Marshal.GetLastWin32Error()); }
        return handle;
    }

    internal static bool TokenFlag(int kind)
    {
        Check(OpenProcessToken(Process.GetCurrentProcess().Handle, 8, out var token));
        using (token) { Check(GetTokenInformation(token, kind, out var value, 4, out _)); return value != 0; }
    }

    internal static async Task<Dictionary<string, string>> Run(string exe, string arguments, string policy, IntPtr sid, IntPtr canary)
    {
        var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Inherit = true };
        Check(CreatePipe(out var stdoutRead, out var stdoutWrite, ref security, 0));
        using (stdoutRead) using (stdoutWrite)
        {
            Check(CreatePipe(out var stderrRead, out var stderrWrite, ref security, 0));
            using (stderrRead) using (stderrWrite) using (var stdin = OpenInheritableFile("NUL"))
            {
                Check(SetHandleInformation(stdoutRead, 1, 0)); Check(SetHandleInformation(stderrRead, 1, 0));
                var attributes = IntPtr.Zero;
                var initialized = false;
                var allocated = new List<IntPtr>();
                SafeAccessTokenHandle? restricted = null;
                ProcessInformation info = default;
                Process? process = null;
                WorkerJob? job = null;
                try
                {
                    nuint size = 0;
                    var count = policy == "appcontainer" ? 3 : 2;
                    _ = InitializeProcThreadAttributeList(IntPtr.Zero, count, 0, ref size);
                    attributes = Marshal.AllocHGlobal((nint)size);
                    Check(InitializeProcThreadAttributeList(attributes, count, 0, ref size));
                    initialized = true;
                    void Attribute<T>(nuint key, T value) where T : struct
                    {
                        var pointer = Marshal.AllocHGlobal(Marshal.SizeOf<T>()); allocated.Add(pointer);
                        Marshal.StructureToPtr(value, pointer, false);
                        Check(UpdateProcThreadAttribute(attributes, 0, key, pointer, (nuint)Marshal.SizeOf<T>(), IntPtr.Zero, IntPtr.Zero));
                    }
                    if (policy == "handle-control")
                        Attribute(0x20002, new CanaryHandleList { Input = stdin.DangerousGetHandle(), Output = stdoutWrite.DangerousGetHandle(), Error = stderrWrite.DangerousGetHandle(), Canary = canary });
                    else Attribute(0x20002, new HandleList { Input = stdin.DangerousGetHandle(), Output = stdoutWrite.DangerousGetHandle(), Error = stderrWrite.DangerousGetHandle() });
                    if (policy == "appcontainer") Attribute(0x20009, new SecurityCapabilities { Sid = sid });
                    job = new WorkerJob(@"Global\QuickExcel.Probe." + Guid.NewGuid().ToString("N"));
                    Attribute(0x2000D, job.Handle);
                    if (policy == "stripped-token")
                    {
                        Check(OpenProcessToken(Process.GetCurrentProcess().Handle, 0xF01FF, out var token));
                        using (token) Check(CreateRestrictedToken(token, 1 | 4, 0, IntPtr.Zero, 0, IntPtr.Zero, 0, IntPtr.Zero, out restricted));
                    }
                    var startup = new StartupInfoEx
                    {
                        Info = new StartupInfo { Size = Marshal.SizeOf<StartupInfoEx>(), Flags = 0x100,
                            Input = stdin.DangerousGetHandle(), Output = stdoutWrite.DangerousGetHandle(), Error = stderrWrite.DangerousGetHandle() },
                        Attributes = attributes
                    };
                    const uint flags = 0x80000 | 0x4 | 0x8000000; // extended startup, suspended, no window
                    var command = new StringBuilder($"\"{exe}\" {arguments}");
                    Check(restricted is null
                        ? CreateProcess(exe, command, IntPtr.Zero, IntPtr.Zero, true, flags, IntPtr.Zero, Path.GetDirectoryName(exe)!, ref startup, out info)
                        : CreateProcessAsUser(restricted, exe, command, IntPtr.Zero, IntPtr.Zero, true, flags, IntPtr.Zero, Path.GetDirectoryName(exe)!, ref startup, out info));
                    process = Process.GetProcessById(info.ProcessId);
                    job.VerifyMembership(process.SafeHandle);
                    if (ResumeThread(info.Thread) == uint.MaxValue) throw new Win32Exception(Marshal.GetLastWin32Error());
                    stdoutWrite.Dispose(); stderrWrite.Dispose();
                    using var output = new StreamReader(new FileStream(stdoutRead, FileAccess.Read));
                    using var error = new StreamReader(new FileStream(stderrRead, FileAccess.Read));
                    var stdout = output.ReadToEndAsync(); var stderr = error.ReadToEndAsync();
                    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                    await process.WaitForExitAsync(deadline.Token);
                    var text = await stdout.WaitAsync(deadline.Token); var errors = await stderr.WaitAsync(deadline.Token);
                    if (process.ExitCode != 0) throw new InvalidOperationException($"Probe exit {process.ExitCode:X8}: {errors[..Math.Min(2000, errors.Length)]}");
                    return JsonSerializer.Deserialize<Dictionary<string, string>>(text) ?? throw new InvalidDataException("No probe results.");
                }
                finally
                {
                    job?.Dispose();
                    if (info.Process != IntPtr.Zero) { _ = TerminateProcess(info.Process, 1); CloseHandle(info.Process); }
                    if (info.Thread != IntPtr.Zero) CloseHandle(info.Thread);
                    process?.Dispose(); restricted?.Dispose();
                    if (attributes != IntPtr.Zero) { if (initialized) DeleteProcThreadAttributeList(attributes); Marshal.FreeHGlobal(attributes); }
                    foreach (var pointer in allocated) Marshal.FreeHGlobal(pointer);
                }
            }
        }
    }

    // Ordinary-process positive control only; the production launcher assigns the Job at process creation.
    internal static WorkerJob AttachControlJob(Process process)
    {
        var job = new WorkerJob(@"Global\QuickExcel.Probe." + Guid.NewGuid().ToString("N"));
        try { Check(AssignProcessToJobObject(job.Handle, process.SafeHandle)); return job; }
        catch { job.Dispose(); throw; }
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AssignProcessToJobObject(IntPtr job, SafeProcessHandle process);

    [StructLayout(LayoutKind.Sequential)] private struct SecurityAttributes { public int Length; public IntPtr Descriptor; [MarshalAs(UnmanagedType.Bool)] public bool Inherit; }
    [StructLayout(LayoutKind.Sequential)] private struct HandleList { public IntPtr Input, Output, Error; }
    [StructLayout(LayoutKind.Sequential)] private struct CanaryHandleList { public IntPtr Input, Output, Error, Canary; }
    [StructLayout(LayoutKind.Sequential)] private struct SecurityCapabilities { public IntPtr Sid, Capabilities; public uint CapabilityCount, Reserved; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct StartupInfo
    {
        public int Size; public string? Reserved, Desktop, Title;
        public uint X, Y, XSize, YSize, XChars, YChars, Fill, Flags;
        public ushort Show, ReservedSize; public IntPtr ReservedBytes, Input, Output, Error;
    }
    [StructLayout(LayoutKind.Sequential)] private struct StartupInfoEx { public StartupInfo Info; public IntPtr Attributes; }
    [StructLayout(LayoutKind.Sequential)] private struct ProcessInformation { public IntPtr Process, Thread; public int ProcessId, ThreadId; }

    [DllImport("userenv.dll", CharSet = CharSet.Unicode)] internal static extern int CreateAppContainerProfile(string name, string display, string description, IntPtr capabilities, uint count, out IntPtr sid);
    [DllImport("userenv.dll", CharSet = CharSet.Unicode)] internal static extern int DeleteAppContainerProfile(string name);
    [DllImport("advapi32.dll")] internal static extern IntPtr FreeSid(IntPtr sid);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool OpenProcessToken(IntPtr process, uint access, out SafeAccessTokenHandle token);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool GetTokenInformation(SafeAccessTokenHandle token, int kind, out int value, int size, out int needed);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool CreateRestrictedToken(SafeAccessTokenHandle token, uint flags, uint disableCount, IntPtr disable, uint privilegeCount, IntPtr privileges, uint restrictCount, IntPtr restrict, out SafeAccessTokenHandle result);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateFileW")] private static extern SafeFileHandle CreateFile(string path, uint access, uint share, ref SecurityAttributes security, uint creation, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)] internal static extern bool ReadFile(IntPtr handle, byte[] buffer, int count, out int read, IntPtr overlapped);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetFilePointerEx(SafeFileHandle handle, long distance, out long position, uint method);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool CreatePipe(out SafeFileHandle read, out SafeFileHandle write, ref SecurityAttributes security, uint size);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetHandleInformation(SafeFileHandle handle, uint mask, uint flags);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool InitializeProcThreadAttributeList(IntPtr list, int count, uint flags, ref nuint size);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool UpdateProcThreadAttribute(IntPtr list, uint flags, nuint key, IntPtr value, nuint size, IntPtr previous, IntPtr returned);
    [DllImport("kernel32.dll")] private static extern void DeleteProcThreadAttributeList(IntPtr list);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateProcessW")] private static extern bool CreateProcess(string application, StringBuilder command, IntPtr processAttributes, IntPtr threadAttributes, bool inherit, uint flags, IntPtr environment, string directory, ref StartupInfoEx startup, out ProcessInformation info);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateProcessAsUserW")] private static extern bool CreateProcessAsUser(SafeAccessTokenHandle token, string application, StringBuilder command, IntPtr processAttributes, IntPtr threadAttributes, bool inherit, uint flags, IntPtr environment, string directory, ref StartupInfoEx startup, out ProcessInformation info);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll")] private static extern bool TerminateProcess(IntPtr process, uint exit);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
}
