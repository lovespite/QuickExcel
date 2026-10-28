using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using QuickExcel.ScriptContracts;

internal static class IsolationChecks
{
    internal static string Run(FunctionCall call)
    {
        var path = (string)call.Arguments[0].ToObject()!;
        var handle = new IntPtr(long.Parse((string)call.Arguments[1].ToObject()!));
        var results = new Dictionary<string, bool>();
        bool Denied(Action action)
        {
            try { action(); return false; }
            catch (UnauthorizedAccessException) { return true; }
        }
        results["privateReadDenied"] = Denied(() => File.ReadAllText(path));
        results["privateWriteDenied"] = Denied(() => File.WriteAllText(path, "unexpected"));
        results["payloadWriteDenied"] = Denied(() => File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "unexpected"), "unexpected"));
        var ownership = Directory.GetParent(AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar))!.FullName;
        results["ownershipReadDenied"] = Denied(() => Directory.GetFileSystemEntries(ownership));
        results["ownershipWriteDenied"] = Denied(() => File.WriteAllText(Path.Combine(ownership, "unexpected"), "unexpected"));
        var buffer = new byte[128];
        try
        {
            results["canaryNotInherited"] = GetFileType(handle) != 1 || !ReadFile(handle, buffer, buffer.Length, out var read, IntPtr.Zero) ||
                Encoding.UTF8.GetString(buffer, 0, read) != "QuickExcel private canary";
        }
        // Windows strict-handle checking may raise a native exception for an absent handle instead of returning ERROR_INVALID_HANDLE.
        catch (SEHException) { results["canaryNotInherited"] = true; }
        results["environmentClean"] = Environment.GetEnvironmentVariable("QUICKEXCEL_TEST_SECRET") is null &&
            Environment.GetEnvironmentVariable("HTTP_PROXY") is null && Environment.GetEnvironmentVariable("DOTNET_STARTUP_HOOKS") is null &&
            Environment.GetEnvironmentVariable("DOTNET_EnableDiagnostics") == "0";
        results["childProcessDenied"] = false;
        try
        {
            using var child = Process.Start(new ProcessStartInfo(Environment.ProcessPath!) { UseShellExecute = false, CreateNoWindow = true })!;
            child.Kill(); child.WaitForExit();
        }
        catch (Win32Exception e) { results["childProcessDenied"] = e.NativeErrorCode == 1816; }
        return JsonSerializer.Serialize(results);
    }
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadFile(IntPtr handle, byte[] buffer, int count, out int read, IntPtr overlapped);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint GetFileType(IntPtr handle);
}
