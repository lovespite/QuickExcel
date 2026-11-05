using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>One command path for grid edits, formula-bar edits, paste, clear, structure and history.</summary>
public sealed class EditorController(WorkbookDocument document, Func<DataValidationException, bool>? confirmValidation = null)
{
    public CommandResult SetCell(int sheetIndex, int row, int column, string value) =>
        SetCells(sheetIndex, [(row, column, value)]);

    public CommandResult SetCells(int sheetIndex, IEnumerable<(int Row, int Column, string Value)> values)
    {
        var edits = values.Take(20_001).ToArray();
        if (edits.Length > 20_000) throw new ArgumentException(UiText.Get("单次编辑上限为 20,000 个单元格。"));
        if (edits.Any(edit => edit.Row < 0 || edit.Row >= WorkbookLimits.MaxRows || edit.Column < 0 || edit.Column >= WorkbookLimits.MaxColumns))
            throw new ArgumentException(UiText.Format($"界面编辑范围为 {WorkbookLimits.MaxRows:N0} 行 × {WorkbookLimits.MaxColumns:N0} 列。"));
        var snapshot = document.Snapshot; var sheet = snapshot.Sheets[sheetIndex];
        return Execute(new SetCellsCommand(edits.Select(v => new CellEdit(new(sheet.Id, v.Row, v.Column),
            CellInput.ParseForCell(v.Value, snapshot, sheet, v.Row, v.Column)))), snapshot.Revision);
    }

    public CommandResult ClearCells(int sheetIndex, IEnumerable<(int Row, int Column)> positions)
    {
        var cells = positions.Take(20_001).ToArray();
        if (cells.Length > 20_000) throw new ArgumentException(UiText.Get("单次编辑上限为 20,000 个单元格。"));
        if (cells.Any(cell => cell.Row < 0 || cell.Row >= WorkbookLimits.MaxRows || cell.Column < 0 || cell.Column >= WorkbookLimits.MaxColumns))
            throw new ArgumentException(UiText.Format($"界面编辑范围为 {WorkbookLimits.MaxRows:N0} 行 × {WorkbookLimits.MaxColumns:N0} 列。"));
        var snapshot = document.Snapshot;
        var sheet = snapshot.Sheets[sheetIndex];
        return Execute(new ClearCellsCommand(cells.Select(cell => new CellAddress(sheet.Id, cell.Row, cell.Column))), snapshot.Revision);
    }

    public WorksheetSnapshot AddSheet(string name) => document.AddSheet(name);
    public CommandResult Execute(DocumentCommand command, long? expectedRevision = null)
    {
        var revision = expectedRevision ?? document.Session.Revision;
        try { return document.Session.Execute(command, revision); }
        catch (DataValidationException ex) when (ex.CanOverride)
        {
            if (confirmValidation?.Invoke(ex) != true) throw;
            return document.Session.Execute(command, revision, acceptValidationWarnings: true);
        }
    }
    public async Task<CommandResult> ExecuteAsync(DocumentCommand command,
        Func<DataValidationException, Task<bool>> confirm, long? expectedRevision = null)
    {
        var revision = expectedRevision ?? document.Session.Revision;
        try { return document.Session.Execute(command, revision); }
        catch (DataValidationException ex) when (ex.CanOverride)
        {
            if (!await confirm(ex)) throw;
            return document.Session.Execute(command, revision, acceptValidationWarnings: true);
        }
    }
    public void UpdateScriptSource(string source) => document.ScriptSource = source;
    public CommandResult Undo() => document.Session.UndoWithResult();
    public CommandResult Redo() => document.Session.RedoWithResult();
}
