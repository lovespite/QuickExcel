using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Win32;

namespace QuickExcel.Desktop.Avalonia;

internal static class AssociationSmoke
{
    private const string Associations = @"Software\QuickExcel\Capabilities\FileAssociations";
    private const string Classes = @"Software\Classes\";

    internal static int Run(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>();
        Exception? failure = null;
        try
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("Windows associations require Windows.");
            CheckRegistration(directory, (passed, label) =>
            {
                if (!passed) throw new InvalidOperationException(label);
                checks.Add(label);
            });
        }
        catch (Exception ex) { failure = ex; }
        using var output = File.Create(Path.Combine(directory, "result.json"));
        using var json = new Utf8JsonWriter(output, new JsonWriterOptions { Indented = true });
        json.WriteStartObject();
        json.WriteBoolean("passed", failure is null);
        json.WriteStartArray("checks");
        foreach (var label in checks) json.WriteStringValue(label);
        json.WriteEndArray();
        json.WriteString("error", failure?.ToString());
        json.WriteBoolean("realUserAssociationsChanged", false);
        json.WriteEndObject();
        return failure is null ? 0 : 1;
    }

    [SupportedOSPlatform("windows")]
    internal static void CheckRegistration(string directory, Action<bool, string> check)
    {
        var name = @"Software\QuickExcel-Association-Test-" + Guid.NewGuid().ToString("N");
        var payloadRoot = Path.Combine(directory, "association-payloads-" + Guid.NewGuid().ToString("N"));
        var previous = CreatePayload("旧版本 with spaces");
        var current = CreatePayload("新版本 with spaces");
        try
        {
            using var root = Registry.CurrentUser.CreateSubKey(name);
            using (var fresh = root.CreateSubKey("Fresh"))
            {
                check(WindowsFileAssociations.ReadSelection(fresh).SequenceEqual([".xlsq"]), "Fresh settings offer XLSQ without registering it");
                check(!WindowsFileAssociations.RepairForDeployment(fresh, current, "true", false) && fresh.SubKeyCount == 0,
                    "A first ClickOnce launch does not register types the user never selected");
                WindowsFileAssociations.Register(fresh, current, []);
                var empty = Snapshot(fresh);
                check(!WindowsFileAssociations.RepairExisting(fresh, current) && Snapshot(fresh) == empty,
                    "An explicitly empty selection remains empty after an update");
            }

            using var user = root.CreateSubKey("Existing");
            WindowsFileAssociations.Register(user, previous, WindowsFileAssociations.Extensions);
            var hashes = new HashSet<string>(StringComparer.Ordinal);
            foreach (var ext in WindowsFileAssociations.Extensions)
            {
                var icon = Icon(previous, ext);
                check(Equals(Read(user, Type(ext) + @"\DefaultIcon"), $"\"{icon}\",0"), ext + " registers its own packaged document icon");
                hashes.Add(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(icon))));
                foreach (var size in new[] { 16, 32, 48, 256 })
                {
                    var handle = LoadImage(IntPtr.Zero, icon, 1, size, size, 0x10);
                    try { check(handle != IntPtr.Zero, $"Windows decodes {ext} document icon at {size} pixels"); }
                    finally { if (handle != IntPtr.Zero) DestroyIcon(handle); }
                }
                Write(user, Classes + ext, "", "OtherApp" + ext);
                Write(user, UserChoice(ext), "ProgId", "OtherApp" + ext);
                Write(user, UserChoice(ext), "Hash", "preserve-existing-choice");
                Write(user, OpenWith(ext), "OtherApp" + ext, Array.Empty<byte>(), RegistryValueKind.None);
            }
            check(hashes.Count == 3, "XLSQ, XLSX and CSV use distinct icon assets");
            Write(user, @"Software\RegisteredApplications", "OtherApp", @"Software\OtherApp\Capabilities");
            // Match releases before this fix: XLSX and CSV had no DefaultIcon at all.
            user.DeleteSubKeyTree(Type(".xlsx") + @"\DefaultIcon");
            user.DeleteSubKeyTree(Type(".csv") + @"\DefaultIcon");
            var original = Snapshot(user);
            foreach (var deployed in new string?[] { null, "false", "invalid" })
                check(!WindowsFileAssociations.RepairForDeployment(user, current, deployed, false) && Snapshot(user) == original,
                    "Non-ClickOnce startup leaves existing associations unchanged: " + (deployed ?? "unset"));
            check(!WindowsFileAssociations.RepairForDeployment(user, current, "true", true) && Snapshot(user) == original,
                "Isolated self-check, smoke and user-data launches never repair real-user associations");
            check(WindowsFileAssociations.RepairForDeployment(user, current, "TrUe", false),
                "A normal ClickOnce launch repairs an existing registration after its cache path changes");
            foreach (var ext in WindowsFileAssociations.Extensions)
            {
                CheckTarget(user, ext, current, check);
                check(Equals(Read(user, Classes + ext), "OtherApp" + ext) &&
                    Equals(Read(user, UserChoice(ext), "ProgId"), "OtherApp" + ext) &&
                    Equals(Read(user, UserChoice(ext), "Hash"), "preserve-existing-choice") &&
                    Read(user, OpenWith(ext), "OtherApp" + ext) is byte[],
                    ext + " repair preserves defaults, UserChoice hash and other OpenWith applications");
            }
            check(Equals(Read(user, @"Software\RegisteredApplications", "OtherApp"), @"Software\OtherApp\Capabilities"),
                "Repair preserves other registered applications");
            var repaired = Snapshot(user);
            check(!WindowsFileAssociations.RepairExisting(user, current) && Snapshot(user) == repaired,
                "A second launch with current paths is an idempotent no-op");

            WindowsFileAssociations.Register(user, current, [".xlsx"]);
            check(WindowsFileAssociations.ReadSelection(user).SequenceEqual([".xlsx"]), "Registration persists the exact selected types");
            check(WindowsFileAssociations.RepairExisting(user, previous), "An update also refreshes retained, deselected ProgIDs still usable by Windows defaults");
            foreach (var ext in new[] { ".xlsq", ".csv" })
            {
                CheckTarget(user, ext, previous, check);
                check(Read(user, Associations, ext) is null && Read(user, OpenWith(ext), "QuickExcel" + ext) is null,
                    ext + " remains unadvertised after its retained ProgID is repaired");
            }
            user.DeleteSubKeyTree(Type(".xlsq"));
            check(WindowsFileAssociations.RepairExisting(user, current) && !Exists(user, Type(".xlsq")),
                "A removed, deselected ProgID is not recreated");
            user.DeleteSubKeyTree(Type(".xlsx"));
            using (var openWith = user.OpenSubKey(OpenWith(".xlsx"), writable: true)) openWith!.DeleteValue("QuickExcel.xlsx");
            using (var programs = user.OpenSubKey(@"Software\RegisteredApplications", writable: true)) programs!.DeleteValue("QuickExcel");
            check(WindowsFileAssociations.RepairExisting(user, current), "Persisted selection restores missing registration records");
            CheckTarget(user, ".xlsx", current, check);
            check(Read(user, OpenWith(".xlsx"), "QuickExcel.xlsx") is byte[] &&
                Equals(Read(user, @"Software\RegisteredApplications", "QuickExcel"), @"Software\QuickExcel\Capabilities"),
                "Repair restores selected OpenWith and registered application records");
            Write(user, Associations, ".csv", "OtherApp.csv");
            Write(user, Associations, ".tsv", "OtherApp.tsv");
            check(WindowsFileAssociations.ReadSelection(user).SequenceEqual([".xlsx"]), "Foreign capability values are not treated as selected QuickExcel types");
            WindowsFileAssociations.RepairExisting(user, previous);
            check(Equals(Read(user, Associations, ".csv"), "OtherApp.csv") && Equals(Read(user, Associations, ".tsv"), "OtherApp.tsv"),
                "Repair preserves foreign and unsupported capability mappings");

            var incomplete = CreatePayload("incomplete update");
            var missingIcon = Icon(incomplete, ".csv");
            File.Delete(missingIcon);
            var beforeFailure = Snapshot(user);
            check(MissingFile(() => WindowsFileAssociations.RepairExisting(user, incomplete)) && Snapshot(user) == beforeFailure,
                "A missing required document icon rejects the update before any registry mutation");
            File.Copy(Icon(current, ".csv"), missingIcon);
            check(WindowsFileAssociations.RepairExisting(user, incomplete), "The next launch retries successfully after the payload is repaired");
            CheckTarget(user, ".xlsx", incomplete, check);
            CheckTarget(user, ".csv", incomplete, check);
            var missingExe = Path.Combine(payloadRoot, "missing", "quickxl.exe");
            var beforeMissingExe = Snapshot(user);
            check(MissingFile(() => WindowsFileAssociations.RepairExisting(user, missingExe)) && Snapshot(user) == beforeMissingExe,
                "A missing executable cannot replace working association paths");
        }
        finally { Registry.CurrentUser.DeleteSubKeyTree(name, throwOnMissingSubKey: false); }

        string CreatePayload(string folder)
        {
            var output = Path.Combine(payloadRoot, folder);
            Directory.CreateDirectory(Path.Combine(output, "Assets"));
            var executable = Path.Combine(output, "quickxl.exe");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "quickxl.exe"), executable);
            foreach (var ext in WindowsFileAssociations.Extensions)
                File.Copy(Icon(Path.Combine(AppContext.BaseDirectory, "quickxl.exe"), ext), Icon(executable, ext));
            return executable;
        }
    }

    private static string Icon(string executable, string ext) => Path.Combine(Path.GetDirectoryName(executable)!, "Assets", "QuickExcel-" + ext[1..].ToUpperInvariant() + ".ico");
    private static string Type(string ext) => Classes + "QuickExcel" + ext;
    private static string OpenWith(string ext) => Classes + ext + @"\OpenWithProgids";
    private static string UserChoice(string ext) => @"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" + ext + @"\UserChoice";

    [SupportedOSPlatform("windows")]
    private static void CheckTarget(RegistryKey user, string ext, string executable, Action<bool, string> check) =>
        check(Equals(Read(user, Type(ext) + @"\shell\open\command"), $"\"{executable}\" \"%1\"") &&
            Equals(Read(user, Type(ext) + @"\DefaultIcon"), $"\"{Icon(executable, ext)}\",0"),
            ext + " command and document icon point to the current payload with quoted paths");

    [SupportedOSPlatform("windows")]
    private static object? Read(RegistryKey user, string path, string value = "")
    { using var key = user.OpenSubKey(path); return key?.GetValue(value); }

    [SupportedOSPlatform("windows")]
    private static bool Exists(RegistryKey user, string path)
    { using var key = user.OpenSubKey(path); return key is not null; }

    [SupportedOSPlatform("windows")]
    private static void Write(RegistryKey user, string path, string name, object value, RegistryValueKind kind = RegistryValueKind.String)
    { using var key = user.CreateSubKey(path); key.SetValue(name, value, kind); }

    [SupportedOSPlatform("windows")]
    private static string Snapshot(RegistryKey root)
    {
        var values = new List<string>();
        Visit(root, "");
        return string.Join("\n", values);
        void Visit(RegistryKey key, string path)
        {
            values.Add(path);
            foreach (var name in key.GetValueNames().Order(StringComparer.Ordinal))
            {
                var value = key.GetValue(name);
                values.Add(path + "|" + name + "|" + key.GetValueKind(name) + "|" + (value is byte[] bytes ? Convert.ToBase64String(bytes) : value));
            }
            foreach (var name in key.GetSubKeyNames().Order(StringComparer.Ordinal))
            { using var child = key.OpenSubKey(name); Visit(child!, path + "\\" + name); }
        }
    }

    private static bool MissingFile(Action action)
    { try { action(); return false; } catch (FileNotFoundException) { return true; } }

    [DllImport("user32.dll", EntryPoint = "LoadImageW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadImage(IntPtr instance, string name, uint type, int width, int height, uint flags);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr icon);
}
