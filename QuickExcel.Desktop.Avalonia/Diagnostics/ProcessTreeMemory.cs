using System.Diagnostics;
using System.Runtime.InteropServices;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Windows private commit for the host and every observed descendant, including browser/renderer/utility processes.</summary>
internal sealed class ProcessTreeMemory : IAsyncDisposable
{
    internal sealed record ProcessMemory(int Id, int Parent, string Name, long Bytes);
    internal sealed record Sample(double Milliseconds, long Bytes, int[] Processes, ProcessMemory[] Details);
    private readonly List<Sample> _samples = [];
    private readonly CancellationTokenSource _stop = new();
    private readonly Stopwatch _watch = Stopwatch.StartNew();
    private readonly Task _sampling;
    internal IReadOnlyList<Sample> Samples { get { lock (_samples) return _samples.ToArray(); } }
    internal ProcessTreeMemory() { if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("此基准的 private commit 端点仅定义于 Windows。"); _sampling = Task.Run(SampleAsync); }
    private async Task SampleAsync()
    {
        // Retain handles until exit so an observed PID cannot be recycled into an unrelated
        // process between samples; also reject older children of a recycled parent PID.
        var root = Process.GetCurrentProcess(); _ = root.SafeHandle;
        var known = new Dictionary<int, Process> { [Environment.ProcessId] = root };
        try
        {
            while (!_stop.IsCancellationRequested)
            {
                var entries = new List<(int Id, int Parent)>(); var snapshot = CreateToolhelp32Snapshot(2, 0);
                if (snapshot == new nint(-1)) throw new IOException("无法枚举浏览器进程树。");
                try
                {
                    var entry = new ProcessEntry { Size = (uint)Marshal.SizeOf<ProcessEntry>(), Name = "" };
                    if (Process32FirstW(snapshot, ref entry)) do { entries.Add(((int)entry.ProcessId, (int)entry.ParentProcessId)); } while (Process32NextW(snapshot, ref entry));
                }
                finally { CloseHandle(snapshot); }
                foreach (var pair in known.ToArray()) if (pair.Value.HasExited) { pair.Value.Dispose(); known.Remove(pair.Key); }
                bool added;
                do
                {
                    added = false;
                    foreach (var item in entries)
                    {
                        if (known.ContainsKey(item.Id) || !known.TryGetValue(item.Parent, out var parent)) continue;
                        Process? child = null;
                        try
                        {
                            child = Process.GetProcessById(item.Id); _ = child.SafeHandle;
                            if (child.StartTime.ToUniversalTime() < parent.StartTime.ToUniversalTime()) continue;
                            known.Add(item.Id, child); child = null; added = true;
                        }
                        catch (ArgumentException) { } catch (InvalidOperationException) { }
                        finally { child?.Dispose(); }
                    }
                } while (added);
                long bytes = 0; var alive = new List<int>(); var details = new List<ProcessMemory>();
                foreach (var (id, process) in known)
                {
                    try { process.Refresh(); var size = process.PrivateMemorySize64; bytes += size; alive.Add(id); details.Add(new(id, entries.First(e => e.Id == id).Parent, process.ProcessName, size)); }
                    catch (ArgumentException) { } catch (InvalidOperationException) { }
                }
                lock (_samples) _samples.Add(new(_watch.Elapsed.TotalMilliseconds, bytes, alive.ToArray(), details.ToArray()));
                await Task.Delay(100, _stop.Token);
            }
        }
        catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
        finally { foreach (var process in known.Values) process.Dispose(); }
    }
    public async ValueTask DisposeAsync() { await _stop.CancelAsync(); await _sampling; _stop.Dispose(); }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct ProcessEntry
    { public uint Size, Usage, ProcessId; public nuint DefaultHeapId; public uint ModuleId, Threads, ParentProcessId; public int BasePriority; public uint Flags; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string Name; }
    [DllImport("kernel32.dll", SetLastError = true)] private static extern nint CreateToolhelp32Snapshot(uint flags, uint process);
    [DllImport("kernel32.dll", ExactSpelling = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool Process32FirstW(nint snapshot, ref ProcessEntry entry);
    [DllImport("kernel32.dll", ExactSpelling = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool Process32NextW(nint snapshot, ref ProcessEntry entry);
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CloseHandle(nint handle);
}
