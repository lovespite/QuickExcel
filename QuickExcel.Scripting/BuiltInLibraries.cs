using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Scripting;

/// <summary>Trusted, version-pinned package resources. They are never accepted from workbook IPC.</summary>
internal static class BuiltInLibraries
{
    private static readonly (NativeModule[] Modules, string Support) Payload = Read();
    internal static NativeModule[] Modules => Payload.Modules;
    internal static string Support => Payload.Support;

    private static (NativeModule[], string) Read()
    {
        using var runtimeStream = typeof(BuiltInLibraries).Assembly.GetManifestResourceStream("QuickExcel.ScriptLibraries.Runtime")!;
        using var catalogStream = typeof(BuiltInLibraries).Assembly.GetManifestResourceStream("QuickExcel.ScriptLibraries.Catalog")!;
        using var runtime = JsonDocument.Parse(runtimeStream);
        using var catalog = JsonDocument.Parse(catalogStream);
        var expected = catalog.RootElement;
        // The canonical compact catalog is produced by the build tool and bound
        // into the worker handshake, compilation cache and trust policy.
        var compact = JsonSerializer.Serialize(expected, ScriptingJsonContext.Default.JsonElement);
        if (Hash(compact) != BuiltInLibraryIdentity.Value) throw new InvalidOperationException("Built-in catalog identity mismatch.");
        var modules = new List<NativeModule>();
        foreach (var property in runtime.RootElement.GetProperty("modules").EnumerateObject())
        {
            var source = property.Value.GetString()!;
            if (Hash(source) != expected.GetProperty("packages").GetProperty(property.Name).GetProperty("sha256").GetString())
                throw new InvalidOperationException("Built-in module integrity mismatch.");
            modules.Add(new(property.Name, source, 0));
        }
        var support = runtime.RootElement.GetProperty("support").GetString()!;
        if (Hash(support) != expected.GetProperty("supportSha256").GetString())
            throw new InvalidOperationException("Built-in adapter integrity mismatch.");
        return (modules.ToArray(), support);
    }
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
}
