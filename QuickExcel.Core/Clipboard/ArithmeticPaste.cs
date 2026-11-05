using System.Globalization;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public enum PasteArithmetic { Add, Subtract, Multiply, Divide }

/// <summary>Numeric value operands; formulas at the destination retain their expression and dependencies.</summary>
public sealed class ArithmeticPasteCommand(ExternalClipboardRegion source, SheetId sheetId, CellRectangle selection,
    long expectedRevision, PasteArithmetic operation, bool skipBlanks = true) : DocumentCommand
{
    public static CellRectangle OutputRange(ExternalClipboardRegion source, CellRectangle selection)
    {
        WorkbookClipboard.ValidateRange(selection);
        var scalar = source.Rows == 1 && source.Columns == 1;
        if (!scalar && selection.Count != 1 && (selection.Rows != source.Rows || selection.Columns != source.Columns))
            throw new ArgumentException(MessageText.Get("选区尺寸须与粘贴数据相同，或只选择左上角单元格。"));
        var output = scalar ? selection : new(selection.Top, selection.Left, checked(selection.Top + source.Rows - 1), checked(selection.Left + source.Columns - 1));
        WorkbookClipboard.ValidateRange(output); return output;
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (snapshot.Revision != expectedRevision) throw new InvalidOperationException(MessageText.Get("文档已改变，请重新预览。"));
        if (!Enum.IsDefined(operation)) throw new ArgumentException(MessageText.Get("未知运算粘贴类型。"));
        source.RequireMode(CellPasteMode.Values);
        if (source.ValuesAreDisplayText || !source.Merges.IsEmpty || source.HasFilterMetadata)
            throw new NotSupportedException(MessageText.Get("运算粘贴需要无合并、无筛选元数据的原始数值，不能使用 HTML 显示文本。"));
        var output = OutputRange(source, selection); var sheet = snapshot.GetSheet(sheetId);
        WorkbookClipboard.RequirePlainGeometry(sheet, output); WorkbookClipboard.RequireVisibleDimensions(sheet, output);
        var hidden = WorksheetFilterEngine.Evaluate(snapshot, sheetId);
        if (hidden.Any(row => row >= output.Top && row <= output.Bottom))
            throw new InvalidOperationException(MessageText.Get("运算粘贴目标包含筛除行，请先清除筛选或调整选区。"));
        var scalar = source.Rows == 1 && source.Columns == 1;
        var edits = new List<CellEdit>();
        var symbol = operation switch { PasteArithmetic.Add => "+", PasteArithmetic.Subtract => "-", PasteArithmetic.Multiply => "*", _ => "/" };
        for (var row = 0; row < output.Rows; row++)
        for (var column = 0; column < output.Columns; column++)
        {
            var operand = source.Cells[scalar ? 0 : row * source.Columns + column];
            if (skipBlanks && operand.Value.Kind == CellValueKind.Blank && operand.Formula is null) continue;
            var r = output.Top + row; var c = output.Left + column; var current = sheet.GetCell(r, c);
            if (current?.SpillOwner is not null || current?.Content.DynamicArray == true)
                throw new InvalidOperationException(MessageText.Get("运算粘贴不能修改数组公式或溢出结果。"));
            if (current?.Content.Kind == CellContentKind.Formula && (current.Calculation?.Value is null || current.Calculation.InputRevision != snapshot.Revision ||
                current.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)))
                throw new InvalidOperationException(MessageText.Get("运算粘贴目标公式需要最新计算结果。"));
            double Number(CellValue value)
            {
                if (value.Kind == CellValueKind.Blank) return 0;
                if (value.Kind != CellValueKind.Number) throw new InvalidOperationException(MessageText.Get("运算粘贴只接受数字或真正的空白；文本、布尔和错误会使整批取消。"));
                return value.NumberValue;
            }
            var left = Number(current?.Value ?? CellValue.Blank); var right = Number(operand.Value);
            if (operation == PasteArithmetic.Divide && right == 0) throw new InvalidOperationException(MessageText.Get("运算粘贴遇到除零，整批未修改。"));
            var result = operation switch { PasteArithmetic.Add => left + right, PasteArithmetic.Subtract => left - right, PasteArithmetic.Multiply => left * right, _ => left / right };
            if (!double.IsFinite(result)) throw new InvalidOperationException(MessageText.Get("运算粘贴结果溢出，整批未修改。"));
            var style = sheet.GetStyleId(r, c);
            CellContent content;
            if (current?.Content.Kind == CellContentKind.Formula)
            {
                var formula = "=(" + current.Content.FormulaText.TrimStart('=') + ")" + symbol + "(" + right.ToString("R", CultureInfo.InvariantCulture) + ")";
                _ = FormulaSyntax.Parse(formula); content = CellContent.FromFormula(formula, style);
            }
            else content = CellContent.FromLiteral(CellValue.FromNumber(result), style);
            edits.Add(new(new(sheetId, r, c), content));
        }
        return new SetCellsCommand(edits).Prepare(snapshot);
    }
}
