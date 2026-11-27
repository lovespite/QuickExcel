namespace QuickExcel.Core;

public sealed record ScriptFileFilter(string Name, IReadOnlyList<string> Extensions);
public sealed record ScriptFilePickerRequest(string Action, bool Directory, bool Save, bool Multiple,
    string Access, string Mode, bool AllowDelete, string? Title, string? SuggestedName, IReadOnlyList<ScriptFileFilter> Filters);
public sealed record ScriptFileSelection(IReadOnlyList<string> LocalPaths, string? WorkbookPath = null);

/// <summary>Trusted desktop boundary. Local paths never enter the worker protocol.</summary>
public delegate Task<ScriptFileSelection> ScriptFilePicker(ScriptFilePickerRequest request,
    Func<IDisposable> beginUserInteraction, CancellationToken cancellation);

