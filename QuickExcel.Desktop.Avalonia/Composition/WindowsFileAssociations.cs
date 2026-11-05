using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Win32;

namespace QuickExcel.Desktop.Avalonia;

[SupportedOSPlatform("windows")]
internal static class WindowsFileAssociations
{
    internal static readonly string[] Extensions = [".xlsq", ".xlsx", ".csv"];
    private const string Capabilities = @"Software\QuickExcel\Capabilities";
    private const string Associations = Capabilities + @"\FileAssociations";

    internal static string[] ReadSelection(RegistryKey user)
    {
        using var key = user.OpenSubKey(Associations);
        return key is null ? [".xlsq"] : Extensions.Where(ext => key.GetValue(ext) is string).ToArray();
    }

    internal static void Register(RegistryKey user, string executable, IReadOnlyCollection<string> selected)
    {
        using var capabilities = user.CreateSubKey(Capabilities);
        capabilities.SetValue("ApplicationName", "QuickExcel");
        capabilities.SetValue("ApplicationDescription", "QuickExcel 工作簿与 CSV 文件");
        using var associations = user.CreateSubKey(Associations);
        foreach (var ext in Extensions)
        {
            var progId = "QuickExcel" + ext;
            var openWithPath = @"Software\Classes\" + ext + @"\OpenWithProgids";
            if (!selected.Contains(ext))
            {
                associations.DeleteValue(ext, false);
                using var previous = user.OpenSubKey(openWithPath, writable: true);
                previous?.DeleteValue(progId, false);
                continue;
            }
            using var type = user.CreateSubKey(@"Software\Classes\" + progId);
            type.SetValue("", $"QuickExcel {ext[1..].ToUpperInvariant()} 文件");
            using var command = type.CreateSubKey(@"shell\open\command");
            command.SetValue("", $"\"{executable}\" \"%1\"");
            using var openWith = user.CreateSubKey(openWithPath);
            openWith.SetValue(progId, Array.Empty<byte>(), RegistryValueKind.None);
            associations.SetValue(ext, progId);
        }
        using var programs = user.CreateSubKey(@"Software\RegisteredApplications");
        programs.SetValue("QuickExcel", Capabilities);
    }

    internal static void ApplyAndOpenDefaults(IReadOnlyCollection<string> selected)
    {
        var executable = Path.Combine(AppContext.BaseDirectory, "QuickExcel.Avalonia.exe");
        if (!File.Exists(executable)) throw new FileNotFoundException("找不到 QuickExcel 启动程序。", executable);
        Register(Registry.CurrentUser, executable, selected);
        SHChangeNotify(0x08000000, 0, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED
        Process.Start(new ProcessStartInfo("ms-settings:defaultapps?registeredAppUser=QuickExcel") { UseShellExecute = true });
    }

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(uint eventId, uint flags, IntPtr item1, IntPtr item2);
}
