using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace QuickExcel.Runtime;

/// <summary>Resource/lifetime containment only. This does not restrict file or network access.</summary>
internal sealed class WorkerJob : IDisposable
{
    private readonly SafeFileHandle _handle;
    public WorkerJob(Process process)
    {
        _handle = CreateJobObject(IntPtr.Zero, null);
        if (_handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        try
        {
            var limits = new ExtendedLimits
            {
                Basic = new BasicLimits { Flags = 0x2000 | 0x100 | 0x8 | 0x400, ActiveProcesses = 1 },
                ProcessMemory = 256 * 1024 * 1024
            };
            if (!SetInformationJobObject(_handle, 9, ref limits, (uint)Marshal.SizeOf<ExtendedLimits>()) ||
                !AssignProcessToJobObject(_handle, process.SafeHandle))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        catch { _handle.Dispose(); throw; }
    }
    public void Dispose() => _handle.Dispose();

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
    private static extern SafeFileHandle CreateJobObject(IntPtr attributes, string? name);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(SafeFileHandle job, int infoClass, ref ExtendedLimits info, uint length);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(SafeFileHandle job, SafeProcessHandle process);
}
