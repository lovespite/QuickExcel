using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed partial class ScriptWorkbookModel
{
    private ScriptStorage _sessionStorage = ScriptStorage.Empty;
    private bool _localStorageWritten;
    internal bool HasSessionStorageWrites { get; private set; }
    internal ScriptStorage SessionStorageResult => _sessionStorage;
    internal void InitializeSessionStorage(ScriptStorage storage) => _sessionStorage = storage;
    public ScriptStorage ReadStorage(bool session)
    {
        _check();
        return session ? _sessionStorage : _working.LocalStorage;
    }
    public void WriteStorage(bool session, string operation, string? key = null, string? value = null)
    {
        try
        {
            _check();
            if (_invalid) throw new InvalidOperationException("本次脚本已有无效写入，不能提交。");
            if (_working.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(MessageText.Get(WorkbookFileState.FormulaReadOnlyMessage));
            var storage = session ? _sessionStorage : _working.LocalStorage;
            var next = operation switch
            {
                "set" => storage.Set(key!, value!, session ? ScriptStorage.MaxSessionBytes : null),
                "remove" => storage.Remove(key!),
                "clear" => ScriptStorage.Empty,
                _ => throw new ArgumentException("未知存储操作。")
            };
            if (session) { _sessionStorage = next; HasSessionStorageWrites = true; }
            else
            {
                _localStorageWritten = true;
                if (!ReferenceEquals(storage, next)) _working = _working.WithLocalStorage(next);
            }
        }
        catch { _invalid = true; throw; }
    }
}
