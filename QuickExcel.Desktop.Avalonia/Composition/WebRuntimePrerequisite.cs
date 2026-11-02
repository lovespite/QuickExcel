using System.Runtime.Versioning;
using Microsoft.Win32;

namespace QuickExcel.Desktop.Avalonia;

internal static class WebRuntimePrerequisite
{
    internal const string DownloadUrl = "https://developer.microsoft.com/microsoft-edge/webview2/";
    internal static string? Version => OperatingSystem.IsWindows() ? WindowsInstallation()?.Version : null;
    private sealed record Installation(string Version, string Directory);
    [SupportedOSPlatform("windows")]
    private static Installation? WindowsInstallation()
    {
        const string key = @"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
        {
            using var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry32);
            using var registration = root.OpenSubKey(key);
            using var state = root.OpenSubKey(key.Replace(@"\Clients\", @"\ClientState\"));
            if (registration?.GetValue("pv") is string version && System.Version.TryParse(version, out var parsed) && parsed.Major > 0 &&
                state?.GetValue("EBWebView") is string directory && Path.IsPathFullyQualified(directory) &&
                File.Exists(Path.Combine(directory, "msedgewebview2.exe")) && File.Exists(Path.Combine(directory, "EBWebView", "x64", "EmbeddedBrowserWebView.dll"))) return new(version, directory);
        }
        return null;
    }
    internal static string? Require()
    {
        // Presence is a deployment preflight. The ready/worker handshake still checks
        // actual browser capabilities and never silently substitutes another editor.
        if (!OperatingSystem.IsWindows()) return null;
        var installation = WindowsInstallation();
        if (installation is null)
            throw new InvalidOperationException("未检测到 Microsoft Edge WebView2 Runtime。请安装 Evergreen Runtime（离线设备使用 x64 Standalone Installer），然后重新打开此页面。下载：" + DownloadUrl);
        return installation.Directory;
    }
}
