using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;

namespace QuickExcel.Runtime;

/// <summary>Resource/lifetime containment only. This does not restrict file or network access.</summary>
internal sealed class WorkerJob : IDisposable
{
    private readonly SafeFileHandle _handle;
    internal IntPtr Handle => _handle.DangerousGetHandle();
    public WorkerJob(string name)
    {
#pragma warning disable CA1416
        using var identity = WindowsIdentity.GetCurrent();
        var descriptor = new RawSecurityDescriptor($"O:{identity.User!.Value}D:P(A;;GA;;;{identity.User.Value})");
        var bytes = new byte[descriptor.BinaryLength]; descriptor.GetBinaryForm(bytes, 0);
#pragma warning restore CA1416
        var pointer = Marshal.AllocHGlobal(bytes.Length);
        int error;
        try
        {
            Marshal.Copy(bytes, 0, pointer, bytes.Length);
            var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Descriptor = pointer };
            _handle = CreateJobObject(ref security, name);
            error = Marshal.GetLastWin32Error();
        }
        finally { Marshal.FreeHGlobal(pointer); }
        try
        {
            if (_handle.IsInvalid || error == 183) throw new Win32Exception(error); // Never adopt an existing named object.
            var limits = new ExtendedLimits
            {
                Basic = new BasicLimits { Flags = 0x2000 | 0x100 | 0x8 | 0x400, ActiveProcesses = 1 },
                ProcessMemory = 256 * 1024 * 1024
            };
            if (!SetInformationJobObject(_handle, 9, ref limits, (uint)Marshal.SizeOf<ExtendedLimits>()))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        catch { _handle.Dispose(); throw; }
    }
    public void Dispose() => _handle.Dispose();

    internal void VerifyMembership(SafeProcessHandle process)
    {
        if (!IsProcessInJob(process, _handle, out var contained)) throw new Win32Exception(Marshal.GetLastWin32Error());
        if (!contained) throw new IOException("Worker was not created in its containment Job.");
    }

    internal static void EnsureEmpty(string name)
    {
        using var job = OpenJobObject(4, false, name); // Query only; do not kill an unrelated process by PID.
        if (job.IsInvalid)
        {
            var error = Marshal.GetLastWin32Error();
            if (error == 2) return; // Job objects survive while associated processes still run.
            throw new Win32Exception(error);
        }
        var timer = Stopwatch.StartNew();
        do
        {
            if (!QueryInformationJobObject(job, 1, out var info, (uint)Marshal.SizeOf<BasicAccounting>(), IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            if (info.ActiveProcesses == 0) return;
            Thread.Sleep(20);
        } while (timer.Elapsed < TimeSpan.FromSeconds(1));
        throw new IOException("Worker Job still has active processes; sandbox recovery deferred.");
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SecurityAttributes { public int Length; public IntPtr Descriptor; [MarshalAs(UnmanagedType.Bool)] public bool Inherit; }
    [StructLayout(LayoutKind.Sequential)]
    private struct BasicAccounting
    {
        public long UserTime, KernelTime, PeriodUserTime, PeriodKernelTime;
        public uint PageFaults, TotalProcesses, ActiveProcesses, TerminatedProcesses;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BasicLimits
    {
        public long ProcessTime, JobTime;
        public uint Flags;
        public nuint MinimumWorkingSet, MaximumWorkingSet;
        public uint ActiveProcesses;
        public nuint Affinity;
        public uint Priority, Scheduling;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct ExtendedLimits
    {
        public BasicLimits Basic;
        public ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes;
        public nuint ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
    }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateJobObjectW")]
    private static extern SafeFileHandle CreateJobObject(ref SecurityAttributes attributes, string name);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "OpenJobObjectW")]
    private static extern SafeFileHandle OpenJobObject(uint access, bool inherit, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool QueryInformationJobObject(SafeFileHandle job, int kind, out BasicAccounting info, uint size, IntPtr returned);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(SafeFileHandle job, int infoClass, ref ExtendedLimits info, uint length);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsProcessInJob(SafeProcessHandle process, SafeFileHandle job, out bool result);
}
