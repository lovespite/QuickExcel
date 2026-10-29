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

    public CommandResult ClearCells(int sheetIndex, IEnumerable<(int Row, int Column)> positions) =>
        SetCells(sheetIndex, positions.Select(position => (position.Row, position.Column, "")));

    public WorksheetSnapshot AddSheet(string name) => document.AddSheet(name);
    public CommandResult Execute(DocumentCommand command) => document.Session.Execute(command);
    public void UpdateScriptSource(string source) => document.ScriptSource = source;
    public CommandResult Undo() => document.Session.UndoWithResult();
    public CommandResult Redo() => document.Session.RedoWithResult();
}
