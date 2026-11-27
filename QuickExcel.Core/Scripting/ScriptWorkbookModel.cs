using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Host-owned, command-backed view of an action. Reads observe its prior writes;
/// only DocumentSession can publish the complete transaction.</summary>
public sealed partial class ScriptWorkbookModel
{
    private readonly WorkbookSnapshot _original;
    private WorkbookSnapshot _working;
    private readonly Action _check;
    private readonly Func<WorkbookSnapshot, CancellationToken, CalculationResult>? _calculate;
    private WorkbookSnapshot? _calculatedView;
    private bool? _initialResultsPending;
    private ImmutableList<DocumentCommand> _commands = [];
    private long _writes;
    private long _reads;
    private bool _invalid;
    private WorkbookDocument? _fileDraft;
    private readonly HashSet<int> _reservedTableIds;
    internal ScriptWorkbookModel(WorkbookSnapshot snapshot, Action check, ImmutableArray<CellRectangle> selection,
        Func<WorkbookSnapshot, CancellationToken, CalculationResult>? calculate, HashSet<int> reservedTableIds)
    { _original = _working = snapshot; _check = check; Selection = selection; _calculate = calculate; _reservedTableIds = reservedTableIds; }

    /// <summary>Expose the ordinary script model over an isolated, script-disabled file draft.
    /// Each command goes through the private document session's validation and revision tracking.</summary>
    public static ScriptWorkbookModel ForFileDraft(WorkbookDocument document, Action check)
    {
        void Check()
        {
            check();
            if (document.ScriptsEnabled) throw new InvalidOperationException("Selected workbook scripts must be disabled.");
        }
        Check();
        return new ScriptWorkbookModel(document.Snapshot, Check, [], null, []) { _fileDraft = document };
    }
    public int ReserveTableId()
    {
        _check();
        lock (_reservedTableIds)
        {
            _reservedTableIds.UnionWith(_working.Sheets.SelectMany(s => s.Data.Tables).Select(t => t.Id));
            var id = checked(_reservedTableIds.DefaultIfEmpty().Max() + 1); _reservedTableIds.Add(id); return id;
        }
    }
    public ImmutableArray<CellRectangle> Selection { get; }
    public bool HasPendingCalculation { get; private set; }
    public WorkbookSnapshot View { get { _check(); return _working; } }
    public WorkbookSnapshot ValueView { get { _check(); return _calculatedView ?? _working; } }
    internal bool HasWrites => _commands.Count > 0 || _localStorageWritten;

    public CellSnapshot? ReadCell(SheetId sheet, CellPosition position, bool original = false, bool calculated = false)
    {
        _check(); WorkbookLimits.ValidatePosition(position.Row, position.Column);
        if (++_reads > 100_000) throw new InvalidOperationException("每次脚本最多读取 100000 个单元格。");
        return (original ? _original : calculated ? _calculatedView ?? _working : _working).GetSheet(sheet).GetCell(position.Row, position.Column);
    }

    /// <summary>Calculate privately before a value read. Commands and legacy snapshots remain input-only;
    /// provisional results never publish events, history or changes to the real document.</summary>
    public void PrepareValueRead(SheetId sheetId, CellRectangle range, CancellationToken cancellationToken)
    {
        _check(); cancellationToken.ThrowIfCancellationRequested();
        if (_calculatedView is not null) return;
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (range.Bottom < range.Top || range.Right < range.Left || range.Count > 100_000 - _reads)
            throw new InvalidOperationException("每次脚本最多读取 100000 个单元格。");
        if (_fileDraft is not null)
        {
            // Imported caches (including XLSQ caches marked fresh by a previous session)
            // cannot establish a value in a file opened with scripts disabled.
            if (!_fileDraft.Calculate(cancellationToken)) throw new InvalidOperationException("Selected workbook calculation was superseded.");
            _working = _calculatedView = _fileDraft.Snapshot;
            return;
        }
        var sheet = _working.GetSheet(sheetId);
        var needed = false;
        foreach (var address in Addresses(sheetId, range))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = sheet.GetCell(address.Row, address.Column);
            var formula = cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null;
            var blank = cell is null || !formula && cell.Content.LiteralValue.Kind == CellValueKind.Blank;
            // A blank may become a newly created/grown array's spill child.
            if (formula && (HasPendingCalculation || cell!.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.Value is null) ||
                blank && (HasPendingCalculation || InitialResultsPending(cancellationToken)))
            { needed = true; break; }
        }
        if (!needed) return;
        try
        {
            if (_calculate is null) throw new InvalidOperationException("此会话未配置事务内公式计算器。");
            var result = _calculate(_working, cancellationToken);
            _check(); cancellationToken.ThrowIfCancellationRequested();
            var calculated = DocumentSession.ApplyCalculation(_working, result, out _);
            _check(); cancellationToken.ThrowIfCancellationRequested();
            _calculatedView = calculated;
            HasPendingCalculation = false;
        }
        catch { _invalid = true; throw; }
    }

    private bool InitialResultsPending(CancellationToken cancellationToken)
    {
        if (_initialResultsPending is { } pending) return pending;
        foreach (var sheet in _original.Sheets)
        foreach (var cell in sheet.Cells.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if ((cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null) &&
                (cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.Value is null))
                return (_initialResultsPending = true).Value;
        }
        return (_initialResultsPending = false).Value;
    }

    public void SetValues(SheetId sheetId, CellRectangle range, IReadOnlyList<IReadOnlyList<CellValue>> values)
        => Write(sheetId, range, () =>
        {
            if (values.Count != range.Bottom - range.Top + 1 || values.Any(r => r.Count != range.Right - range.Left + 1))
                throw new ArgumentException("值矩阵必须与区域大小一致。");
            var sheet = _working.GetSheet(sheetId);
            return new SetScriptCellsCommand(Addresses(sheetId, range).Select(a => new CellEdit(a,
                CellContent.FromLiteral(values[a.Row - range.Top][a.Column - range.Left], sheet.GetStyleId(a.Row, a.Column)))));
        });

    public void SetFormulas(SheetId sheetId, CellRectangle range, IReadOnlyList<IReadOnlyList<string>> formulas)
        => Write(sheetId, range, () =>
        {
            if (formulas.Count != range.Bottom - range.Top + 1 || formulas.Any(r => r.Count != range.Right - range.Left + 1))
                throw new ArgumentException("公式矩阵必须与区域大小一致。");
            foreach (var formula in formulas.SelectMany(r => r))
            {
                if (formula is null || formula.Length is < 2 or > 8192 || formula[0] != '=') throw new ArgumentException("公式必须以 = 开头且不超过 8192 字符。");
                _ = FormulaSyntax.Parse(formula);
            }
            var sheet = _working.GetSheet(sheetId);
            return new SetScriptCellsCommand(Addresses(sheetId, range).Select(a => new CellEdit(a,
                CellContent.FromFormula(formulas[a.Row - range.Top][a.Column - range.Left], sheet.GetStyleId(a.Row, a.Column)))));
        });

    public void SetFill(SheetId sheetId, CellRectangle range, string? color)
        => Write(sheetId, range, () =>
        {
            if (color is not null && (color.Length != 7 || color[0] != '#' || color.AsSpan(1).ContainsAnyExcept("0123456789abcdefABCDEF")))
                throw new ArgumentException("填充颜色使用 #RRGGBB，例如 #ff0000。");
            return new SetCellStyleCommand(Addresses(sheetId, range), color is null ? new(ClearFillColor: true)
                : new(FillColor: "FF" + color[1..].ToUpperInvariant()));
        });

    /// <summary>Apply a validated adapter command to this private transaction, with the same
    /// protection, cumulative cell budget and rollback rules as value writes.</summary>
    public void ExecuteCommand(SheetId sheetId, CellRectangle affected, DocumentCommand command)
        => Write(sheetId, affected, () =>
        {
            if (command is SetAutoFilterCommand or SetTableFilterCommand)
            {
                PrepareValueRead(sheetId, affected, CancellationToken.None);
                foreach (var filter in _working.GetSheet(sheetId).Data.AllFilters)
                    PrepareValueRead(sheetId, filter.Range, CancellationToken.None);
                _ = command.Prepare(ValueView); // Validate current values once, without publishing or replaying UDFs.
                return command is SetAutoFilterCommand worksheetFilter ? worksheetFilter.WithValidatedValues() : ((SetTableFilterCommand)command).WithValidatedValues();
            }
            if (command is SortRangeCommand sort)
            {
                // Freeze only the evaluated keys. Commit reuses them instead of
                // running UDFs a second time or sorting against stale caches.
                PrepareValueRead(sheetId, affected, CancellationToken.None);
                var descriptor = sort.Descriptor;
                var keys = new Dictionary<CellPosition, CellValue>();
                for (var row = affected.Top + (descriptor.HasHeader ? 1 : 0); row <= affected.Bottom; row++)
                foreach (var key in descriptor.Keys)
                {
                    var position = new CellPosition(row, key.Column);
                    var cell = ReadCell(sheetId, position, calculated: true);
                    if ((cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null) &&
                        (cell!.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.Value is null))
                        throw new InvalidOperationException("排序键无法获得当前结果。");
                    keys[position] = cell?.Value ?? CellValue.Blank;
                }
                return sort.WithKeyValues(keys);
            }
            return command;
        });

    public void ExecuteObjectCommand(SheetId sheetId, int affectedCells, DocumentCommand command)
    {
        if (affectedCells is < 1 or > 20_000) { _invalid = true; throw new InvalidOperationException("对象修改超过 20000 格。"); }
        Write(sheetId, new(0, 0, 0, 0), () => command, affectedCells);
    }

    private void Write(SheetId sheetId, CellRectangle range, Func<DocumentCommand> create, int? chargedCells = null)
    {
        try
        {
            _check();
            if (_invalid) throw new InvalidOperationException("本次脚本已有无效写入，不能提交。");
            WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
            if (range.Bottom < range.Top || range.Right < range.Left ||
                (long)(range.Bottom - range.Top + 1) * (range.Right - range.Left + 1) > 20_000 || _commands.Count >= 4096)
                throw new InvalidOperationException("每次脚本最多修改 20000 格、提交 4096 个模型命令。");
            _writes += chargedCells ?? range.Count;
            if (_writes > 20_000) throw new InvalidOperationException("每次脚本累计最多修改 20000 格（重复写入也计数）。");
            if (_working.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(MessageText.Get(WorkbookFileState.FormulaReadOnlyMessage));
            var command = create();
            if (_fileDraft is not null)
            {
                if (_fileDraft.Session.Execute(command, _working.Revision).Changed)
                {
                    _working = _fileDraft.Snapshot;
                    _calculatedView = null;
                }
                return;
            }
            var draft = command.Prepare(_working);
            ExtendedObjectSafety.ValidateEdit(_working, draft.Snapshot, command, draft.StructureChanged);
            PivotSafety.ValidateEdit(_working, draft.Snapshot, command, draft.StructureChanged);
            FormulaSpillSafety.ValidateEdit(_working, draft.Snapshot, draft.StructureChanged);
            ProtectionSafety.Validate(_working, draft.Snapshot, command);
            TableSafety.ValidateEdit(_working, draft.Snapshot);
            if (!draft.CalculationUnchanged && !ReferenceEquals(_working, draft.Snapshot))
            {
                HasPendingCalculation = true;
                _calculatedView = null;
            }
            else if (_calculatedView is not null && !ReferenceEquals(_working, draft.Snapshot))
            {
                // Keep presentation/protection/filter state in step with the
                // private values without recalculating or publishing the draft.
                _calculatedView = command.Prepare(_calculatedView).Snapshot;
            }
            _commands = _commands.Add(command); _working = draft.Snapshot;
        }
        catch { _invalid = true; throw; }
    }

    /// <summary>Transport validation failures are also fatal to the transaction, even if JS catches them.</summary>
    public void RejectWrite() => _invalid = true;

    // Prepare() replays the immutable command list and local storage as of this call.
    // The clone stays private to saveAs and still checks the live Action's authorization.
    internal ScriptWorkbookModel CaptureForSaveAs()
    {
        _check();
        if (_invalid) throw new InvalidOperationException("本次脚本包含无效写入，不能另存。");
        return (ScriptWorkbookModel)MemberwiseClone();
    }

    internal CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        _check();
        if (_invalid) throw new InvalidOperationException("本次脚本包含无效写入，全部修改已回滚。");
        var current = snapshot;
        var changes = new Dictionary<CellAddress, CellChange>();
        var structure = false; var calculationUnchanged = true;
        foreach (var command in _commands)
        {
            _check();
            var draft = command.Prepare(current);
            ProtectionSafety.Validate(current, draft.Snapshot, command);
            foreach (var change in draft.Changes)
                changes[change.Address] = changes.TryGetValue(change.Address, out var prior) ? change with { Before = prior.Before } : change;
            current = draft.Snapshot; structure |= draft.StructureChanged; calculationUnchanged &= draft.CalculationUnchanged;
        }
        if (_localStorageWritten && !current.LocalStorage.ContentEquals(_working.LocalStorage)) current = current.WithLocalStorage(_working.LocalStorage);
        return new(current, changes.Values.Where(c => c.Before != c.After).ToImmutableArray(), structure, CalculationUnchanged: calculationUnchanged);
    }

    private static IEnumerable<CellAddress> Addresses(SheetId sheet, CellRectangle range)
    {
        for (var row = range.Top; row <= range.Bottom; row++)
        for (var column = range.Left; column <= range.Right; column++) yield return new(sheet, row, column);
    }
}

public sealed partial class DocumentSession
{
    private readonly HashSet<int> _scriptTableIds = [];
    private SheetId? _scriptSelectionSheet;
    private CellPosition _scriptSelectionCell;
    private ImmutableArray<CellRectangle> _scriptSelection = [];
    public void UpdateScriptSelection(SheetId sheet, CellPosition cell, IEnumerable<CellRectangle> ranges)
    {
        lock (_gate)
        {
            ThrowIfClosed(); _ = _snapshot.GetSheet(sheet);
            var selection = ranges.Take(65).ToImmutableArray();
            if (selection.Length is < 1 or > 64) throw new ArgumentException("脚本选区需要 1 至 64 个区域。");
            foreach (var range in selection) { WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right); }
            _scriptSelectionSheet = sheet; _scriptSelectionCell = cell; _scriptSelection = selection;
        }
    }

    private ScriptWorkbookModel CreateScriptModel(ScriptActionRequest request)
    {
        var sheet = request.Snapshot.GetSheet(request.Snapshot.ActiveSheetId);
        var selection = _scriptSelectionSheet == sheet.Id && _scriptSelectionCell == sheet.View.ActiveCell && !_scriptSelection.IsEmpty
            ? _scriptSelection : [sheet.View.Selection?.Range ?? new(sheet.View.ActiveCell.Row, sheet.View.ActiveCell.Column, sheet.View.ActiveCell.Row, sheet.View.ActiveCell.Column)];
        var model = new ScriptWorkbookModel(request.Snapshot, () =>
        {
            lock (_gate)
            {
                ThrowIfClosed(); request.CancellationToken.ThrowIfCancellationRequested();
                if (!Matches(request.Stamp) || _activeActionId != request.OperationId ||
                    (request.ConsoleSource is null ? !_scriptsEnabled : _activeConsoleSource != request.ConsoleSource))
                    throw new InvalidOperationException("脚本模型已过期或未授权。");
            }
        }, selection, _calculator is null ? null : (snapshot, cancellation) => _calculator.Calculate(
            new CalculationRequest(request.Stamp, snapshot, request.ConsoleSource is null)
            { CancellationToken = request.CancellationToken, RuntimeLifetime = request.RuntimeLifetime, IsProvisional = true }, cancellation), _scriptTableIds);
        model.InitializeSessionStorage(_sessionStorage);
        return model;
    }
}
