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
    private const string RegisteredApplications = @"Software\RegisteredApplications";

    private static string ProgId(string ext) => "QuickExcel" + ext;
    private static string TypePath(string ext) => @"Software\Classes\" + ProgId(ext);
    private static string OpenWithPath(string ext) => @"Software\Classes\" + ext + @"\OpenWithProgids";
    internal static string IconPath(string executable, string ext) =>
        Path.Combine(Path.GetDirectoryName(executable)!, "Assets", $"QuickExcel-{ext[1..].ToUpperInvariant()}.ico");

    internal static string[] ReadSelection(RegistryKey user)
    {
        using var key = user.OpenSubKey(Associations);
        return key is null ? [".xlsq"] : Extensions.Where(ext => Equals(key.GetValue(ext), ProgId(ext))).ToArray();
    }

    internal static void Register(RegistryKey user, string executable, IReadOnlyCollection<string> selected)
    {
        ValidatePayload(executable, Extensions.Where(selected.Contains));

        WriteCapabilities(user);
        using var associations = user.CreateSubKey(Associations);
        foreach (var ext in Extensions)
        {
            if (!selected.Contains(ext))
            {
                associations.DeleteValue(ext, false);
                using var previous = user.OpenSubKey(OpenWithPath(ext), writable: true);
                previous?.DeleteValue(ProgId(ext), false);
                continue;
            }
            WriteFileType(user, executable, ext);
            WriteOpenWith(user, ext);
            associations.SetValue(ext, ProgId(ext));
        }
        WriteValue(user, RegisteredApplications, "QuickExcel", Capabilities);
    }

    internal static bool RepairForDeployment(RegistryKey user, string executable, string? isNetworkDeployed, bool isolated) =>
        !isolated && bool.TryParse(isNetworkDeployed, out var deployed) && deployed && RepairExisting(user, executable);

    internal static bool RepairExisting(RegistryKey user, string executable)
    {
        // Read the actual registrations, not the settings UI's default XLSQ selection.
        // Keep stable ProgIDs so Windows UserChoice remains valid across cache paths.
        using var associations = user.OpenSubKey(Associations);
        var advertised = Extensions.Where(ext => Equals(associations?.GetValue(ext), ProgId(ext))).ToArray();
        var existing = Extensions.Where(ext =>
        {
            using var type = user.OpenSubKey(TypePath(ext));
            return advertised.Contains(ext) || type is not null;
        }).ToArray();
        if (existing.Length == 0) return false;
        ValidatePayload(executable, existing);

        var changed = false;
        foreach (var ext in existing)
        {
            changed |= WriteFileType(user, executable, ext);
            // Deselection removes advertisements, but Windows may still use the ProgID.
            // Refresh its paths without re-enabling a deselected file type.
            if (advertised.Contains(ext)) changed |= WriteOpenWith(user, ext);
        }
        if (advertised.Length > 0)
        {
            changed |= WriteCapabilities(user);
            changed |= WriteValue(user, RegisteredApplications, "QuickExcel", Capabilities);
        }
        return changed;
    }

    internal static string? RepairCurrentDeployment(bool isolated)
    {
        try
        {
            if (RepairForDeployment(Registry.CurrentUser, Path.Combine(AppContext.BaseDirectory, "quickxl.exe"),
                Environment.GetEnvironmentVariable("ClickOnce_IsNetworkDeployed"), isolated))
                NotifyShell();
            return null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            // Do not block workbook opening. No success marker is persisted, so the
            // next ClickOnce launch retries any incomplete repair.
            return UiText.Format($"文件关联自动修复失败：{ex.Message}");
        }
    }

    private static void ValidatePayload(string executable, IEnumerable<string> extensions)
    {
        if (!File.Exists(executable)) throw new FileNotFoundException(UiText.Get("找不到 QuickExcel 启动程序。"), executable);
        foreach (var ext in extensions)
        {
            var icon = IconPath(executable, ext);
            if (!File.Exists(icon)) throw new FileNotFoundException(null, icon);
        }
    }

    private static bool WriteCapabilities(RegistryKey user)
    {
        var changed = WriteValue(user, Capabilities, "ApplicationName", "QuickExcel");
        changed |= WriteValue(user, Capabilities, "ApplicationDescription", UiText.Get("QuickExcel 工作簿与 CSV 文件"));
        return changed;
    }

    private static bool WriteFileType(RegistryKey user, string executable, string ext)
    {
        var path = TypePath(ext);
        var changed = WriteValue(user, path, "", UiText.Format($"QuickExcel {ext[1..].ToUpperInvariant()} 文件"));
        changed |= WriteValue(user, path + @"\DefaultIcon", "", $"\"{IconPath(executable, ext)}\",0");
        changed |= WriteValue(user, path + @"\shell\open\command", "", $"\"{executable}\" \"%1\"");
        return changed;
    }

    private static bool WriteValue(RegistryKey user, string path, string name, string value)
    {
        using (var existing = user.OpenSubKey(path))
            if (Equals(existing?.GetValue(name), value)) return false;
        using var key = user.CreateSubKey(path);
        key.SetValue(name, value);
        return true;
    }

    private static bool WriteOpenWith(RegistryKey user, string ext)
    {
        using (var existing = user.OpenSubKey(OpenWithPath(ext)))
            if (existing?.GetValue(ProgId(ext)) is byte[] { Length: 0 } && existing.GetValueKind(ProgId(ext)) == RegistryValueKind.None)
                return false;
        using var key = user.CreateSubKey(OpenWithPath(ext));
        key.SetValue(ProgId(ext), Array.Empty<byte>(), RegistryValueKind.None);
        return true;
    }

    internal static void ApplyAndOpenDefaults(IReadOnlyCollection<string> selected)
    {
        var executable = Path.Combine(AppContext.BaseDirectory, "quickxl.exe");
        if (!File.Exists(executable)) throw new FileNotFoundException(UiText.Get("找不到 QuickExcel 启动程序。"), executable);
        Register(Registry.CurrentUser, executable, selected);
        NotifyShell();
        Process.Start(new ProcessStartInfo("ms-settings:defaultapps?registeredAppUser=QuickExcel") { UseShellExecute = true });
    }

    private static void NotifyShell() => SHChangeNotify(0x08000000, 0, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(uint eventId, uint flags, IntPtr item1, IntPtr item2);
}
