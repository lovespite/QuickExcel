using System.Collections.Immutable;
using System.Text;

namespace QuickExcel.Domain;

public enum ScriptLanguage { JavaScript, TypeScript }
public sealed record ScriptModule(string Path, string Source, ScriptLanguage Language = ScriptLanguage.JavaScript);
public sealed record ScriptBinding(string Module, string Kind);
// Public certificate and signature only; private keys and trust decisions never enter a workbook.
public sealed record ScriptSignature(int Version, string Algorithm, string Certificate, string Value);

/// <summary>Source metadata only. It never carries execution authorization.</summary>
public sealed class ScriptModuleSet
{
    public const string TypeScriptOptions = "ts5.9.3-es2022-esnext-strict-isolated-verbatim-define-v1";
    public ScriptModuleSet(IEnumerable<KeyValuePair<string, ScriptModule>> modules, IEnumerable<string> entryModules,
        long revision = 0, long bindingRevision = 0, int scriptApiVersion = 1, IEnumerable<ScriptBinding>? bindings = null,
        ScriptSignature? signature = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(revision);
        ArgumentOutOfRangeException.ThrowIfNegative(bindingRevision);
        if (scriptApiVersion is not (1 or 2)) throw new ArgumentException("Unsupported script API version.");
        Modules = modules.ToImmutableDictionary(StringComparer.Ordinal);
        EntryModules = entryModules.ToImmutableArray();
        if (Modules.Count is < 1 or > 64 || EntryModules.Length is < 1 or > 64 ||
            EntryModules.Distinct(StringComparer.Ordinal).Count() != EntryModules.Length || EntryModules.Any(id => !Modules.ContainsKey(id)))
            throw new ArgumentException("Invalid script module/entry mapping.");
        Bindings = (bindings ?? EntryModules.Select(id => new ScriptBinding(id, "formula-functions"))).ToImmutableArray();
        if (Bindings.Length is < 1 or > 128 || Bindings.Distinct().Count() != Bindings.Length ||
            Bindings.Any(b => b is null || !EntryModules.Contains(b.Module) || b.Kind is not ("formula-functions" or "actions")) ||
            EntryModules.Any(id => !Bindings.Any(b => b.Module == id)) ||
            scriptApiVersion != 2 && Bindings.Any(b => b.Kind == "actions"))
            throw new ArgumentException("Invalid script bindings or action API version.");
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (id, module) in Modules)
        {
            if (string.IsNullOrEmpty(id) || id.Length > 128 || id is "." or ".." || id.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-' or '.')) ||
                module is null || !Enum.IsDefined(module.Language) || module.Language == ScriptLanguage.TypeScript && scriptApiVersion != 2 ||
                module.Source is null || module.Path is null || module.Path.Length > 512 || module.Path.Any(char.IsControl) || !module.Path.StartsWith("scripts/", StringComparison.Ordinal) ||
                module.Path.Split('/').Any(s => s.Length == 0 || s is "." or "..") ||
                module.Path.IndexOfAny(['\\', ':', '%', '?', '#']) >= 0 || !paths.Add(module.Path))
                throw new ArgumentException("Invalid script module identity or path.");
        }
        Revision = revision;
        BindingRevision = bindingRevision;
        ScriptApiVersion = scriptApiVersion;
        Signature = signature;
    }
    public ImmutableDictionary<string, ScriptModule> Modules { get; }
    public ImmutableArray<string> EntryModules { get; }
    public long Revision { get; }
    public long BindingRevision { get; }
    public int ScriptApiVersion { get; }
    public ImmutableArray<ScriptBinding> Bindings { get; }
    public ScriptSignature? Signature { get; }
    public ScriptModuleSet WithSignature(ScriptSignature? signature) => Signature == signature ? this : new(
        Modules, EntryModules, Revision, BindingRevision, ScriptApiVersion, Bindings, signature);
    public bool HasActions => Bindings.Any(b => b.Kind == "actions");
    public bool HasTypeScript => Modules.Values.Any(m => m.Language == ScriptLanguage.TypeScript);
    public ImmutableArray<string> EntriesFor(string kind) => EntryModules.Where(id => Bindings.Any(b => b.Module == id && b.Kind == kind)).ToImmutableArray();
    public string PrimarySource => Modules[EntryModules[0]].Source;
    public bool CanExecuteLegacy => ScriptApiVersion == 1 && Modules.Count == 1 && EntryModules.Length == 1;
    public bool CanExecute => CanExecuteLegacy || ScriptApiVersion == 2;
    public static ScriptModuleSet Single(string source) => new(
        [new KeyValuePair<string, ScriptModule>("main", new("scripts/main.js", source))], ["main"]);
    public ScriptModuleSet WithPrimarySource(string source) => source == PrimarySource ? this : new(
        Modules.SetItem(EntryModules[0], Modules[EntryModules[0]] with { Source = source }), EntryModules,
        checked(Revision + 1), BindingRevision, ScriptApiVersion, Bindings);
}

public sealed record WorkbookPartBinding(string WorkbookSheetId, string Part);

/// <summary>Immutable import provenance used to prevent silent loss of unsupported file content.</summary>
public sealed class WorkbookFileState
{
    public WorkbookFileState(string sourcePath, int sourceFormatVersion, IEnumerable<byte> originalWorkbook,
        WorkbookSnapshot originalModel, IEnumerable<WorkbookPartBinding> sheetParts, IEnumerable<string> compatibilityIssues,
        string? manifestJson = null, string? stateJson = null, IEnumerable<CompatibilityDiagnostic>? compatibilityDiagnostics = null,
        IEnumerable<PreservedFormula>? preservedFormulas = null)
    {
        SourcePath = Path.GetFullPath(sourcePath);
        SourceFormatVersion = sourceFormatVersion;
        OriginalWorkbook = originalWorkbook.ToImmutableArray();
        OriginalModel = originalModel;
        SheetParts = sheetParts.ToImmutableArray();
        CompatibilityIssues = compatibilityIssues.Distinct().ToImmutableArray();
        CompatibilityDiagnostics = (compatibilityDiagnostics ?? CompatibilityIssues.Select(message => new CompatibilityDiagnostic(
            "xlsx.unmodeled", CompatibilityKind.PreservedOnly, "未建模内容", message))).Distinct().ToImmutableArray();
        ManifestJson = manifestJson;
        StateJson = stateJson;
        PreservedFormulas = preservedFormulas?.ToImmutableArray() ?? [];
    }
    public string SourcePath { get; }
    public int SourceFormatVersion { get; }
    public ImmutableArray<byte> OriginalWorkbook { get; }
    public WorkbookSnapshot OriginalModel { get; }
    public ImmutableArray<WorkbookPartBinding> SheetParts { get; }
    public ImmutableArray<string> CompatibilityIssues { get; }
    public ImmutableArray<CompatibilityDiagnostic> CompatibilityDiagnostics { get; }
    public string? ManifestJson { get; }
    public string? StateJson { get; }
    public ImmutableArray<PreservedFormula> PreservedFormulas { get; }
    public bool FormulaReadOnly => !PreservedFormulas.IsEmpty;
    public const string FormulaReadOnlyMessage = "此工作簿含尚不能安全计算或变换的公式，当前以只读模式打开；可查看源码/来源缓存及另存原包，导航不会写入文件。请在 Excel 中修改，或明确导出已有缓存的值快照。";
    public bool RequiresMigration => SourceFormatVersion == 1;
}
