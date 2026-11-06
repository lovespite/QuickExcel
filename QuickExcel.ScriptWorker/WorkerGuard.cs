using System.Runtime.InteropServices;

/// <summary>Final platform guard, called before any protocol byte or workbook source is read.</summary>
internal static partial class WorkerGuard
{
    internal static void Enter()
    {
        if (OperatingSystem.IsLinux())
        {
            var result = LockLinuxWorker();
            if (result != 0) throw new InvalidOperationException($"Linux worker isolation verification failed ({result}).");
        }
        else if (OperatingSystem.IsMacOS())
        {
            var result = LockMacWorker();
            if (result != 0) throw new InvalidOperationException($"macOS worker isolation verification failed ({result}).");
        }
    }

    [LibraryImport("quickexcel-sandbox", EntryPoint = "qx_linux_lock_worker")]
    private static partial int LockLinuxWorker();

    [LibraryImport("quickexcel-sandbox", EntryPoint = "qx_macos_lock_worker")]
    private static partial int LockMacWorker();
}
