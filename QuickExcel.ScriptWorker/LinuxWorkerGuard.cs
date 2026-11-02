using System.Runtime.InteropServices;

internal static partial class LinuxWorkerGuard
{
    internal static void Enter()
    {
        if (!OperatingSystem.IsLinux()) return;
        var result = LockWorker();
        if (result != 0) throw new InvalidOperationException($"Linux worker isolation verification failed ({result}).");
    }
    [LibraryImport("quickexcel-sandbox", EntryPoint = "qx_linux_lock_worker")]
    private static partial int LockWorker();
}
