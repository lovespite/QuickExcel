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
        quickExcel.functions.register(
            "DISCOUNT", (price, rate) => {
                const result = price * (1 - rate);
                return Math.round(result * 100) / 100;
            }
        );

        // 范围传入数组：=TOTAL(B2:B4)
        quickExcel.functions.register(
            "TOTAL", values => values.reduce(
                (sum, value) => sum + Number(value),
                0
            )
        );
        """;

    private readonly WorkbookServices _services;
    private readonly object _saveGate = new();
    public WorkbookDocument(WorkbookSnapshot snapshot, WorkbookServices services, long historyBudgetBytes = 32 * 1024 * 1024)
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentNullException.ThrowIfNull(services.Storage);
        ArgumentNullException.ThrowIfNull(services.Calculator);
        _services = services;
        Session = new(snapshot, historyBudgetBytes);
    }

    public DocumentSession Session { get; }
    public WorkbookSnapshot Snapshot => Session.Snapshot;
    public bool ScriptsEnabled => Session.ScriptsEnabled;
    public IReadOnlyList<string> ScriptErrors => Session.Diagnostics;
    public bool IsDirty => Session.IsDirty;
    public bool CanUndo => Session.CanUndo;
    public bool CanRedo => Session.CanRedo;
    public string ScriptSource
    {
        get => Snapshot.ScriptSource;
        set => Session.Execute(new UpdateScriptSourceCommand(value));
    }

    public static WorkbookDocument CreateBlank(WorkbookServices services) => new(
        new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(SheetId.New(), "Sheet1")], DefaultScript), services);

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

    public void Save(string path)
    {
        // Capture after acquiring the writer so queued saves cannot overwrite a newer capture with an older one.
        WorkbookSnapshot captured;
        lock (_saveGate)
        {
            captured = Snapshot;
            _services.Storage.Save(path, captured);
        }
        // Notifications run outside the IO gate; a subscriber may request another save.
        Session.MarkSaved(Session.SessionId, captured.Revision);
    }

    public void ExportXlsx(string path)
    {
        lock (_saveGate) _services.Storage.Export(path, Snapshot);
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
            CellInput.Parse(value.Value, sheet.GetCell(value.Row, value.Column)?.Content.StyleId ?? 0)));
        return Session.Execute(new SetCellsCommand(edits), snapshot.Revision);
    }

    public string GetCellInput(int sheetIndex, int row, int column) =>
        Snapshot.Sheets[sheetIndex].GetCell(row, column) is { } cell ? CellInput.EditText(cell.Content) : "";
    public string GetCellValue(int sheetIndex, int row, int column) =>
        CellInput.DisplayText(Snapshot.Sheets[sheetIndex].GetCell(row, column)?.Value ?? CellValue.Blank);
    public bool Undo() => Session.Undo();
    public bool Redo() => Session.Redo();

    public void EnableScriptsAndCalculate()
    {
        Session.SetScriptsEnabled(true);
        Calculate();
    }
    public void DisableScripts() => Session.SetScriptsEnabled(false);
    public void Calculate() => Calculate(CancellationToken.None);
    public bool Calculate(CancellationToken cancellationToken)
    {
        var request = Session.BeginCalculation();
        try
        {
            var result = _services.Calculator.Calculate(request, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return Session.TryCommitCalculation(result);
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
    public void Dispose() => Session.Dispose();
}
