using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    // The file format stores the values. Keep the parameters only while their original
    // coordinates still describe the accepted preview in this document session.
    private sealed record SessionDataTable(CellAddress Formula, CellAddress ColumnInput, double[] Columns,
        CellAddress? RowInput, double[] Rows, CellPosition Output)
    {
        public CellRectangle Bounds => new(Output.Row, Output.Column, Output.Row + Columns.Length,
            Output.Column + (RowInput is null ? 1 : Rows.Length));
        public CleaningPreview Preview(WorkbookSnapshot snapshot, IWorkbookCalculator calculator, CancellationToken token) =>
            DataAnalysis.DataTable(snapshot, calculator, Formula, ColumnInput, Columns, RowInput, Rows, Formula.SheetId, Output, token);
        public bool Overlaps(SessionDataTable other) => Formula.SheetId == other.Formula.SheetId &&
            Bounds.Left <= other.Bounds.Right && other.Bounds.Left <= Bounds.Right &&
            Bounds.Top <= other.Bounds.Bottom && other.Bounds.Top <= Bounds.Bottom;
    }

    private readonly List<SessionDataTable> _dataTables = [];
    private bool _tablesPending, _applyingDataTables;

    private void ObserveDataTableChange(DocumentChange change)
    {
        if (_applyingDataTables) return;
        if (change.StructureChanged) _dataTables.Clear();
        else _dataTables.RemoveAll(t => change.Cells.Any(c => c.Address.SheetId == t.Formula.SheetId &&
            t.Bounds.Contains(c.Address.Row, c.Address.Column)));
        _tablesPending = _dataTables.Count > 0;
    }

    private async Task RefreshDataTablesAsync()
    {
        if (!await CommitEditorsAsync()) return;
        if (_dataTables.Count == 0) { _status.Text = "本次打开期间没有可刷新的模拟运算表。"; return; }
        using var cancellation = new CancellationTokenSource();
        using (BeginExecution("正在刷新模拟运算表…", cancellation))
            await RefreshDataTableValuesAsync(cancellation.Token);
        await CalculatePendingAsync();
        _status.Text = "已刷新模拟运算表。";
    }

    private async Task RefreshDataTableValuesAsync(CancellationToken token, bool automatic = false)
    {
        if (_dataTables.Count == 0) return;
        var document = Document;
        var snapshot = document.Snapshot;
        var tables = _dataTables.ToArray();
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        var edits = await Task.Run(() =>
        {
            var draft = snapshot;
            var changes = new List<CellEdit>();
            foreach (var table in tables)
            {
                var preview = table.Preview(draft, _services.Calculator, timeout.Token);
                changes.AddRange(preview.Changes.Select(c => new CellEdit(c.Address, c.After)));
                draft = preview.Result;
            }
            return changes;
        }, timeout.Token);
        timeout.Token.ThrowIfCancellationRequested();
        if (!ReferenceEquals(document, Document)) throw new InvalidOperationException("工作簿已切换，请重新刷新模拟运算表。");
        _applyingDataTables = true;
        try
        {
            if (edits.Count > 0)
            {
                var command = new SetCellsCommand(edits);
                if (automatic) Document.Session.ApplyAutomaticDataTableResults(command, snapshot.Revision);
                else _workspace.Editor.Execute(command, snapshot.Revision);
            }
            _tablesPending = false;
        }
        finally { _applyingDataTables = false; }
    }
}
