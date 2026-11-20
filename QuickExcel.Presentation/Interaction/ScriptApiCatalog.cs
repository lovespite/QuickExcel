using System.Text.Json;

namespace QuickExcel.Presentation;

public enum ScriptApiKind { Local, Host, EnumValue }

/// <summary>One implemented object member or enum value listed by the capability manifest.</summary>
public sealed record ScriptApiEntry(
    string TypeName,
    string Name,
    ScriptApiKind Kind,
    string? Signature,
    string? HostMethod,
    int? Arity,
    int? Cost,
    int? Value,
    string? Summary)
{
    public string Path => TypeName + "." + Name;
}

/// <summary>UI-free catalog behind the settings API browser.</summary>
public static class ScriptApiCatalog
{
    /// <summary>Reads the embedded capability manifest; unlisted Office Scripts members never appear.</summary>
    public static IReadOnlyList<ScriptApiEntry> Parse(string manifestJson)
    {
        using var document = JsonDocument.Parse(manifestJson);
        var root = document.RootElement;
        var entries = new List<ScriptApiEntry>();
        if (root.TryGetProperty("types", out var types) && types.ValueKind == JsonValueKind.Object)
        {
            foreach (var type in types.EnumerateObject())
            {
                if (!type.Value.TryGetProperty("members", out var members) || members.ValueKind != JsonValueKind.Object) continue;
                foreach (var member in members.EnumerateObject())
                {
                    var value = member.Value;
                    var host = value.TryGetProperty("kind", out var kind) && kind.ValueKind == JsonValueKind.String && kind.GetString() == "host";
                    entries.Add(new ScriptApiEntry(type.Name, member.Name, host ? ScriptApiKind.Host : ScriptApiKind.Local,
                        Text(value, "signature"), Text(value, "host"), Number(value, "arity"), Number(value, "cost"), null, Text(value, "summary")));
                }
            }
        }
        if (root.TryGetProperty("enums", out var enums) && enums.ValueKind == JsonValueKind.Object)
        {
            foreach (var type in enums.EnumerateObject())
            {
                if (type.Value.ValueKind != JsonValueKind.Object) continue;
                foreach (var member in type.Value.EnumerateObject())
                    if (member.Value.ValueKind == JsonValueKind.Number && member.Value.TryGetInt32(out var number))
                        entries.Add(new ScriptApiEntry(type.Name, member.Name, ScriptApiKind.EnumValue, null, null, null, null, number, null));
            }
        }
        return entries.OrderBy(entry => entry.TypeName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public static IReadOnlyList<ScriptApiEntry> Filter(IEnumerable<ScriptApiEntry> entries, string? query)
    {
        var text = query?.Trim();
        if (string.IsNullOrEmpty(text)) return entries as IReadOnlyList<ScriptApiEntry> ?? entries.ToArray();
        return entries.Where(entry =>
            entry.Path.Contains(text, StringComparison.OrdinalIgnoreCase) ||
            entry.Summary?.Contains(text, StringComparison.OrdinalIgnoreCase) == true).ToArray();
    }

    private static string? Text(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() : null;

    private static int? Number(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out var number)
            ? number : null;
}
