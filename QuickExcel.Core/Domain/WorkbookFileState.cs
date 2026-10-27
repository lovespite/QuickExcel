using System.Collections.Immutable;
using System.Text;

namespace QuickExcel.Domain;

public sealed record ScriptModule(string Path, string Source);

/// <summary>Source metadata only. It never carries execution authorization.</summary>
public sealed class ScriptModuleSet
{
    public ScriptModuleSet(IEnumerable<KeyValuePair<string, ScriptModule>> modules, IEnumerable<string> entryModules,
        long revision = 0, long bindingRevision = 0)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(revision);
        ArgumentOutOfRangeException.ThrowIfNegative(bindingRevision);
        Modules = modules.ToImmutableDictionary(StringComparer.Ordinal);
        EntryModules = entryModules.ToImmutableArray();
        if (Modules.Count is < 1 or > 64 || EntryModules.Length is < 1 or > 64 ||
            EntryModules.Distinct(StringComparer.Ordinal).Count() != EntryModules.Length || EntryModules.Any(id => !Modules.ContainsKey(id)))
            throw new ArgumentException("Invalid script module/entry mapping.");
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (id, module) in Modules)
        {
            if (string.IsNullOrEmpty(id) || id.Length > 128 || id.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-' or '.')) ||
                module is null || module.Source is null || !module.Path.StartsWith("scripts/", StringComparison.Ordinal) ||
                module.Path.Split('/').Any(s => s.Length == 0 || s is "." or "..") ||
                module.Path.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 || !paths.Add(module.Path))
                throw new ArgumentException("Invalid script module identity or path.");
        }
        Revision = revision;
        BindingRevision = bindingRevision;
    }
    public ImmutableDictionary<string, ScriptModule> Modules { get; }
    public ImmutableArray<string> EntryModules { get; }
    public long Revision { get; }
    public long BindingRevision { get; }
    public string PrimarySource => Modules[EntryModules[0]].Source;
    public bool CanExecuteLegacy => Modules.Count == 1 && EntryModules.Length == 1;
    public static ScriptModuleSet Single(string source) => new(
        [new KeyValuePair<string, ScriptModule>("main", new("scripts/main.js", source))], ["main"]);
    public ScriptModuleSet WithPrimarySource(string source) => source == PrimarySource ? this : new(
        Modules.SetItem(EntryModules[0], Modules[EntryModules[0]] with { Source = source }), EntryModules,
        checked(Revision + 1), BindingRevision);
}

public sealed record WorkbookPartBinding(string WorkbookSheetId, string Part);

/// <summary>Immutable import provenance used to prevent silent loss of unsupported file content.</summary>
public sealed class WorkbookFileState
{
    public WorkbookFileState(string sourcePath, int sourceFormatVersion, IEnumerable<byte> originalWorkbook,
        WorkbookSnapshot originalModel, IEnumerable<WorkbookPartBinding> sheetParts, IEnumerable<string> compatibilityIssues,
        string? manifestJson = null, string? stateJson = null)
    {
        SourcePath = Path.GetFullPath(sourcePath);
        SourceFormatVersion = sourceFormatVersion;
        OriginalWorkbook = originalWorkbook.ToImmutableArray();
        OriginalModel = originalModel;
        SheetParts = sheetParts.ToImmutableArray();
        CompatibilityIssues = compatibilityIssues.Distinct().ToImmutableArray();
        ManifestJson = manifestJson;
        StateJson = stateJson;
    }
    public string SourcePath { get; }
    public int SourceFormatVersion { get; }
    public ImmutableArray<byte> OriginalWorkbook { get; }
    public WorkbookSnapshot OriginalModel { get; }
    public ImmutableArray<WorkbookPartBinding> SheetParts { get; }
    public ImmutableArray<string> CompatibilityIssues { get; }
    public string? ManifestJson { get; }
    public string? StateJson { get; }
    public bool RequiresMigration => SourceFormatVersion == 1;
}
