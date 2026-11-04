using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace QuickExcel.ScriptContracts;

public sealed record ModuleSource(string Id, string Path, string Source, string Description = "");

/// <summary>API v2 package source, never filenames to open. Digest covers identities, paths, sources and entry order.</summary>
public sealed record LoadModules(ModuleSource[] Modules, string[] EntryModules, string Sha256)
{
    /// <summary>Inert single-line module metadata carried to the action picker; never executed.</summary>
    public const int MaxDescriptionLength = 200;
    public string[] FunctionEntries { get; init; } = [];
    public string[] ActionEntries { get; init; } = [];
    public string ComputeDigest()
    {
        ValidateMapping();
        return Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(new
        {
            apiVersion = 2, entries = EntryModules, functionEntries = FunctionEntries, actionEntries = ActionEntries,
            modules = Modules.OrderBy(m => m.Id, StringComparer.Ordinal).Select(m => new
            { id = m.Id, path = m.Path, sha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(m.Source))), description = m.Description })
        })));
    }

    public void Validate()
    {
        if (ComputeDigest() != Sha256) throw new ProtocolException("Invalid module set digest.");
    }

    private void ValidateMapping()
    {
        if (Modules is null || Modules.Length is < 1 or > 64 || EntryModules is null || EntryModules.Length is < 1 or > 64)
            throw new ProtocolException("Invalid module count.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long sourceBytes = 0;
        foreach (var module in Modules)
        {
            if (module is null || string.IsNullOrEmpty(module.Id) || module.Id.Length > 128 ||
                module.Id is "." or ".." || module.Id.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-' or '.')) || !ids.Add(module.Id) ||
                module.Path is null || module.Path.Length > 512 || !module.Path.StartsWith("scripts/", StringComparison.Ordinal) ||
                module.Path.Any(char.IsControl) || module.Path.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 ||
                module.Path.Split('/').Any(s => s.Length == 0 || s is "." or "..") || !paths.Add(module.Path) || module.Source is null ||
                module.Description is null || module.Description.Length > MaxDescriptionLength || module.Description.Any(char.IsControl))
                throw new ProtocolException("Invalid module identity or path.");
            sourceBytes += Encoding.UTF8.GetByteCount(module.Source);
        }
        if (sourceBytes > 1024 * 1024) throw new ProtocolException("Combined module source exceeds 1 MiB UTF-8.");
        if (EntryModules.Any(id => id is null || !ids.Contains(id)) || EntryModules.Distinct(StringComparer.Ordinal).Count() != EntryModules.Length)
            throw new ProtocolException("Invalid entry module mapping.");
        foreach (var entries in new[] { FunctionEntries, ActionEntries })
            if (entries is null || entries.Length > 64 || entries.Any(id => !EntryModules.Contains(id)) || entries.Distinct(StringComparer.Ordinal).Count() != entries.Length)
                throw new ProtocolException("Invalid bound module mapping.");
    }
}
