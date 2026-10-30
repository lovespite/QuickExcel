using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>One command path for grid edits, formula-bar edits, paste, clear, structure and history.</summary>
internal sealed class EditorController(WorkbookDocument document)
{
    public CommandResult SetCell(int sheetIndex, int row, int column, string value) =>
        SetCells(sheetIndex, [(row, column, value)]);

    public CommandResult SetCells(int sheetIndex, IEnumerable<(int Row, int Column, string Value)> values)
    {
        var edits = values.Take(20_001).ToArray();
        if (edits.Length > 20_000) throw new ArgumentException("单次编辑上限为 20,000 个单元格。");
        if (edits.Any(edit => edit.Row < 0 || edit.Row >= WorkbookLimits.MaxRows || edit.Column < 0 || edit.Column >= 256))
            throw new ArgumentException("界面编辑范围为 100,000 行 × 256 列。");
        return document.SetCells(sheetIndex, edits);
    }

    public CommandResult ClearCells(int sheetIndex, IEnumerable<(int Row, int Column)> positions)
    {
        var cells = positions.Take(20_001).ToArray();
        if (cells.Length > 20_000) throw new ArgumentException("单次编辑上限为 20,000 个单元格。");
        if (cells.Any(cell => cell.Row < 0 || cell.Row >= WorkbookLimits.MaxRows || cell.Column < 0 || cell.Column >= 256))
            throw new ArgumentException("界面编辑范围为 100,000 行 × 256 列。");
        var snapshot = document.Snapshot;
        var sheet = snapshot.Sheets[sheetIndex];
        return document.Session.Execute(new ClearCellsCommand(cells.Select(cell => new CellAddress(sheet.Id, cell.Row, cell.Column))), snapshot.Revision);
    }

    public WorksheetSnapshot AddSheet(string name) => document.AddSheet(name);
    public CommandResult Execute(DocumentCommand command, long? expectedRevision = null) => document.Session.Execute(command, expectedRevision);
    public void UpdateScriptSource(string source) => document.ScriptSource = source;
    public CommandResult Undo() => document.Session.UndoWithResult();
    public CommandResult Redo() => document.Session.RedoWithResult();
}
