using System.Collections.Frozen;
using System.Text.Json;

namespace QuickExcel.ScriptContracts;

/// <summary>Identity of the pinned library code, declarations and data-only adapters.</summary>
public static class BuiltInLibraryIdentity
{
    public static string Value { get; } = Read();
    private static readonly (FrozenDictionary<string,string> Paths, FrozenSet<string> Files) Definitions = ReadDefinitions();
    public static IReadOnlyDictionary<string,string> EntryDefinitions => Definitions.Paths;
    public static bool IsDefinition(string name) => Definitions.Files.Contains(name);
    private static (FrozenDictionary<string,string>,FrozenSet<string>) ReadDefinitions()
    {
        using var stream=typeof(BuiltInLibraryIdentity).Assembly.GetManifestResourceStream("QuickExcel.ScriptLibraries.Definitions")!;
        using var document=JsonDocument.Parse(stream);
        return (document.RootElement.GetProperty("paths").EnumerateObject().ToFrozenDictionary(p=>p.Name,p=>p.Value.GetString()!,StringComparer.Ordinal),
            document.RootElement.GetProperty("files").EnumerateArray().Select(v=>v.GetString()!).ToFrozenSet(StringComparer.Ordinal));
    }
    private static string Read()
    {
        using var stream = typeof(BuiltInLibraryIdentity).Assembly.GetManifestResourceStream("QuickExcel.ScriptLibraries.Identity")
            ?? throw new InvalidOperationException("Built-in library identity is missing.");
        using var reader = new StreamReader(stream);
        var value = reader.ReadToEnd().Trim();
        if (value.Length != 64 || value.Any(c => !char.IsAsciiHexDigitUpper(c)))
            throw new InvalidOperationException("Invalid built-in library identity.");
        return value;
    }
}
