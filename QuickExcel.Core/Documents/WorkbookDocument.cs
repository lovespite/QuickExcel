using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Desktop-facing facade over one document session and injected calculation/storage ports.</summary>
public sealed class WorkbookDocument : IDisposable
{
    public const int MaxRows = WorkbookLimits.MaxRows;
    public const int MaxColumns = WorkbookLimits.MaxColumns;
    public const int MaxCells = WorkbookLimits.MaxCells;
    public const string DefaultScript = """
        // 单元格：=DISCOUNT(B2,D2)*C2
        host.functions.register(
            "DISCOUNT", (price, rate) => {
                const result = price * (1 - rate);
                return Math.round(result * 100) / 100;
            }
        );

        // 范围传入数组：=TOTAL(B2:B4)
        host.functions.register(
            "TOTAL", values => values.reduce(
                (sum, value) => sum + Number(value),
                0
            )
        );
        """;

    private readonly WorkbookServices _services;
    private readonly object _saveGate = new();
    public WorkbookDocument(WorkbookSnapshot snapshot, WorkbookServices services, long historyBudgetBytes = 32 * 1024 * 1024, bool readOnly = false)
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentNullException.ThrowIfNull(services.Storage);
        ArgumentNullException.ThrowIfNull(services.Calculator);
        _services = services;
        Session = new(snapshot, historyBudgetBytes, services.Calculator as IWorkbookValidation, services.Storage as IWorkbookVirtualFiles, services.Calculator, readOnly);
    }

    public DocumentSession Session { get; }
    public WorkbookSnapshot Snapshot => Session.Snapshot;
    public bool ScriptsEnabled => Session.ScriptsEnabled;
    public IReadOnlyList<string> ScriptErrors => Session.Diagnostics;
    public bool IsDirty => Session.IsDirty;
    public bool IsReadOnly => Session.IsReadOnly;
    public bool CanUndo => Session.CanUndo;
    public bool CanRedo => Session.CanRedo;
    public string ScriptSource
    {
        get => Snapshot.ScriptSource;
        set => Session.Execute(new UpdateScriptSourceCommand(value));
    }

    public static WorkbookDocument CreateBlank(WorkbookServices services) => new(
        new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(SheetId.New(), "Sheet1")], textCompatibilityVersion: 2), services);

    public static WorkbookDocument CreateDemo(WorkbookServices services)
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        string[][] rows =
        [
            ["商品", "单价", "数量", "折扣率", "实付金额"],
            ["机械键盘", "299", "2", "0.1", "538.2"],
            ["无线鼠标", "129", "3", "0.05", "367.65"],
            ["显示器支架", "189", "1", "0.15", "160.65"],
            ["合计", "", "", "", "1066.5"],
            [],
            ["试用提示", "先启用 JavaScript，再修改单价或折扣率，查看金额重算。"],
            ["公式示例", "'=DISCOUNT(B2,D2)*C2"]
        ];
        for (var row = 0; row < rows.Length; row++)
        for (var column = 0; column < rows[row].Length; column++)
            cells[new(row, column)] = new(CellInput.Parse(rows[row][column], row == 0 ? 1 : row == 4 ? 2 : 0));
        for (var row = 1; row <= 4; row++)
        {
            var value = cells[new(row, 4)].Value;
            var formula = row == 4 ? "=SUM(E2:E4)" : $"=DISCOUNT(B{row + 1},D{row + 1})*C{row + 1}";
            cells[new(row, 4)] = new(CellContent.FromFormula(formula, row == 4 ? 2 : 0), new(value, CalculationState.Stale, 0, 0, 0));
        }
        var columns = Enumerable.Range(0, 5).ToDictionary(index => index, index => new ColumnMetadata(index == 0 ? 22 : 18));
        var sheet = new WorksheetSnapshot(SheetId.New(), "报价单", cells, columns: columns, tabColor: "FF0F766E");
        return new(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], DefaultScript,
            [new(), new() { Bold = true, FillColor = "FF134E4A", FontColor = "FFFFFFFF" }, new() { Bold = true, FillColor = "FFCCFBF1" }]), services);
    }

    public static WorkbookDocument Open(string path, WorkbookServices services) => new(services.Storage.Open(path), services);

    public void Save(string path, bool discardUnsupported = false)
    {
        Session.RequireWritable();
        // Capture after acquiring the writer so queued saves cannot overwrite a newer capture with an older one.
        WorkbookSnapshot captured;
        lock (_saveGate)
        {
            captured = Snapshot;
            if (_services.Storage is IWorkbookSaveStateStorage stateStorage)
                Session.BindSavedFileState(captured.FileState, stateStorage.SaveWithState(path, captured, discardUnsupported).FileState);
            else if (_services.Storage is IWorkbookStorageOptions options) options.Save(path, captured, discardUnsupported);
            else if (discardUnsupported) throw new NotSupportedException("Storage does not support compatibility choices.");
            else _services.Storage.Save(path, captured);
        }
        // Notifications run outside the IO gate; a subscriber may request another save.
        Session.MarkSaved(Session.SessionId, captured.Revision);
    }

    public void ExportXlsx(string path, bool allowStaleCaches = false, bool discardUnsupported = false)
    {
        lock (_saveGate)
        {
            if (_services.Storage is IWorkbookStorageOptions options) options.Export(path, Snapshot, allowStaleCaches, discardUnsupported);
            else if (allowStaleCaches || discardUnsupported) throw new NotSupportedException("Storage does not support export choices.");
            else _services.Storage.Export(path, Snapshot);
        }
    }

    public FileVersion SaveChecked(string path, FileVersion expected, bool discardUnsupported = false, bool overwriteOriginal = false)
    {
        Session.RequireWritable();
        WorkbookSnapshot captured;
        FileVersion version;
        lock (_saveGate)
        {
            captured = Snapshot;
            if (_services.Storage is not IVersionedWorkbookStorage storage) throw new NotSupportedException(MessageText.Get("存储不支持外部文件冲突检查。"));
            if (overwriteOriginal && !discardUnsupported) throw new ArgumentException(MessageText.Get("覆盖必须明确接受兼容性损失。"), nameof(discardUnsupported));
            if (storage is IWorkbookSaveStateStorage stateStorage)
            {
                var saved = stateStorage.SaveWithState(path, captured, discardUnsupported, expected, overwriteOriginal);
                version = saved.Version;
                Session.BindSavedFileState(captured.FileState, saved.FileState);
            }
            else version = overwriteOriginal ? storage.OverwriteChecked(path, captured, expected)
                : storage.SaveChecked(path, captured, discardUnsupported, expected);
        }
        Session.MarkSaved(Session.SessionId, captured.Revision);
        return version;
    }

    public PortableExportReport InspectPortableExport() => _services.Storage is IPortableWorkbookStorage portable
        ? portable.InspectPortableExport(Snapshot) : throw new NotSupportedException("Storage does not support portable formulas.");

    public void ExportPortableXlsx(string path, bool allowStaleCaches = false, bool discardUnsupported = false)
    {
        lock (_saveGate)
        {
            if (_services.Storage is not IPortableWorkbookStorage portable) throw new NotSupportedException("Storage does not support portable formulas.");
            portable.ExportPortable(path, Snapshot, allowStaleCaches, discardUnsupported);
        }
    }

    public WorksheetSnapshot AddSheet(string? name = null)
    {
        var snapshot = Snapshot;
        if (name is null)
        {
            var number = snapshot.Sheets.Length + 1;
            while (snapshot.Sheets.Any(sheet => string.Equals(sheet.Name, $"Sheet{number}", StringComparison.OrdinalIgnoreCase))) number++;
            name = $"Sheet{number}";
        }
        var command = new AddSheetCommand(name);
        Session.Execute(command, snapshot.Revision);
        return Snapshot.GetSheet(command.Id);
    }

    public void SetCell(int sheetIndex, int row, int column, string text) => SetCells(sheetIndex, [(row, column, text)]);

    public CommandResult SetCells(int sheetIndex, IEnumerable<(int Row, int Column, string Value)> values)
    {
        var snapshot = Snapshot;
        var sheet = snapshot.Sheets[sheetIndex];
        var edits = values.Select(value => new CellEdit(new(sheet.Id, value.Row, value.Column),
            CellInput.ParseForCell(value.Value, snapshot, sheet, value.Row, value.Column)));
        return Session.Execute(new SetCellsCommand(edits), snapshot.Revision);
    }

    public string GetCellInput(int sheetIndex, int row, int column) =>
        CellInput.EditText(Snapshot, Snapshot.Sheets[sheetIndex], row, column);
    public string GetCellValue(int sheetIndex, int row, int column)
    {
        var snapshot = Snapshot;
        var sheet = snapshot.Sheets[sheetIndex];
        return CellDisplayFormatter.Format(sheet.GetCell(row, column)?.Value ?? CellValue.Blank,
            snapshot.Styles[sheet.GetStyleId(row, column)], snapshot.DateSystem);
    }
    public bool Undo() => Session.Undo();
    public bool Redo() => Session.Redo();

    public void EnableScriptsAndCalculate() => EnableScriptsAndCalculate(CancellationToken.None);
    public void EnableScriptsAndCalculate(CancellationToken cancellationToken)
    {
        if (!Snapshot.ScriptModules.CanExecute) throw new NotSupportedException(MessageText.Get(ScriptModuleSet.MigrationMessage));
        Session.SetScriptsEnabled(true);
        if (Snapshot.ScriptModules.HasInitialization)
        {
            try
            {
                _ = ListActions(cancellationToken);
                RunAction("@initialize", cancellationToken);
            }
            catch { Session.InvalidateScriptExecution(); throw; }
        }
        Calculate(cancellationToken);
    }
    public void DisableScripts() => Session.SetScriptsEnabled(false);
    public void Calculate() => Calculate(CancellationToken.None);
    public bool Calculate(CancellationToken cancellationToken) => Calculate(cancellationToken, forceFull: false);
    public bool Calculate(CancellationToken cancellationToken, bool forceFull)
    {
        var request = Session.BeginCalculation() with { ForceFull = forceFull };
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, request.CancellationToken);
        try
        {
            var result = _services.Calculator.Calculate(request, lifetime.Token);
            cancellationToken.ThrowIfCancellationRequested();
            var committed = Session.TryCommitCalculation(result);
            if (!committed) Session.InvalidateScriptRuntime(request.RuntimeLifetime);
            return committed;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested &&
            (request.CancellationToken.IsCancellationRequested || request.RuntimeLifetime.IsCancellationRequested))
        {
            Session.InvalidateScriptRuntime(request.RuntimeLifetime);
            return false; // Superseded/closed session: preserve the existing stale-result contract.
        }
        catch (OperationCanceledException)
        {
            Session.TryFailCalculation(request.Stamp, "计算已取消，保留上次结果。", disableScripts: false);
            throw;
        }
        catch (Exception exception)
        {
            Session.TryFailCalculation(request.Stamp, $"计算未完成：{exception.Message}", disableScripts: request.ScriptsEnabled);
            throw;
        }
    }
    public IReadOnlyList<ScriptActionDescriptor> ListActions(CancellationToken cancellationToken = default)
    {
        var executor = _services.Actions ?? throw new NotSupportedException(MessageText.Get("未配置动作执行器。"));
        var request = Session.BeginAction(null);
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, request.CancellationToken);
        try
        {
            var actions = executor.List(request, lifetime.Token);
            lifetime.Token.ThrowIfCancellationRequested();
            if (!Session.CompleteActionDiscovery(request)) throw new InvalidOperationException(MessageText.Get("动作列表已过期，请重新应用脚本。"));
            _eventModules = request.Snapshot.ScriptModules;
            _eventNames = actions.Where(a => a.Name.StartsWith(ScriptEventNames.Prefix, StringComparison.Ordinal)).Select(a => a.Name).ToHashSet(StringComparer.Ordinal);
            Session.ConfigureScriptEvents(_eventModules, _eventNames.Select(n => n[ScriptEventNames.Prefix.Length..]));
            return actions.Where(a => a.Name != "@initialize" && !a.Name.StartsWith(ScriptEventNames.Prefix, StringComparison.Ordinal)).ToArray();
        }
        catch (Exception e) { Session.InvalidateScriptRuntime(request.RuntimeLifetime); Session.TryFailCalculation(request.Stamp, e.Message, disableScripts: false); throw; }
    }

    public CommandResult RunAction(string name, CancellationToken cancellationToken = default)
        => RunActionRequest(Session.BeginAction(name), cancellationToken);

    /// <summary>Explicit console execution does not load or authorize stored workbook modules.</summary>
    public CommandResult RunConsole(string source, CancellationToken cancellationToken = default)
        => RunActionRequest(Session.BeginConsole(source), cancellationToken);

    private CommandResult RunActionRequest(ScriptActionRequest request, CancellationToken cancellationToken)
    {
        var executor = _services.Actions ?? throw new NotSupportedException(MessageText.Get("未配置动作执行器。"));
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, request.CancellationToken);
        try
        {
            var result = executor.Run(request, lifetime.Token);
            lifetime.Token.ThrowIfCancellationRequested();
            return Session.CommitAction(request, result, lifetime.Token);
        }
        catch (Exception e) { Session.InvalidateScriptRuntime(request.RuntimeLifetime); Session.TryFailCalculation(request.Stamp, e.Message, disableScripts: false, logs: (e as ScriptRuntimeFailureException)?.Logs); throw; }
    }

    private int _dispatchingEvents;
    private ScriptModuleSet? _eventModules;
    private HashSet<string> _eventNames = [];
    public int DispatchScriptEvents(CancellationToken cancellationToken = default)
    {
        if (Interlocked.Exchange(ref _dispatchingEvents, 1) != 0) return 0;
        var count = 0;
        try
        {
            if (Session.ScriptsEnabled && Session.HasScriptEvents)
            {
                if (!ReferenceEquals(_eventModules, Snapshot.ScriptModules)) _ = ListActions(cancellationToken);
                else Session.ConfigureScriptEvents(_eventModules!, _eventNames.Select(n => n[ScriptEventNames.Prefix.Length..]));
            }
            while (count < 256 && Session.ScriptsEnabled && Session.TakeScriptEvent() is { } item)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!ReferenceEquals(_eventModules, Snapshot.ScriptModules)) _ = ListActions(cancellationToken);
                if (!_eventNames.Contains(ScriptEventNames.Prefix + item.Name)) { count++; continue; }
                var request = Session.BeginAction(ScriptEventNames.Prefix + item.Name, item);
                try
                {
                    var result = (_services.Actions ?? throw new NotSupportedException(MessageText.Get("未配置事件执行器。"))).Run(request, cancellationToken);
                    Session.CommitAction(request, result, cancellationToken);
                }
                catch (Exception error)
                {
                    Session.InvalidateScriptRuntime(request.RuntimeLifetime);
                    Session.TryFailCalculation(request.Stamp, $"事件 {item.Name} 失败：{error.Message}", disableScripts: false, logs: (error as ScriptRuntimeFailureException)?.Logs);
                    throw;
                }
                count++;
            }
            return count;
        }
        finally { Volatile.Write(ref _dispatchingEvents, 0); }
    }

    public void NotifyClosed()
    {
        if (Session.IsClosed || !ScriptsEnabled) return;
        Session.QueueCloseEvent();
        try { DispatchScriptEvents(); } catch { /* Closing is final; diagnostics stay on the retiring session. */ }
    }
    public void Dispose() => Session.Dispose();
}
