using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private readonly Dictionary<CellAddress, bool> _nestedSubtotals = [];
    private readonly Dictionary<AutoFilterColumn, Func<CellValue, StyleDefinition, ExcelDateSystem, bool>> _filterPredicates = [];
    private readonly Dictionary<(SheetId, int), bool> _subtotalFiltered = [];
    private Range SubtotalReference(FormulaNode node)
    {
        if (node is ReferenceNode reference) return new(snapshot.Sheets.Single(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase)), reference);
        if (node is FunctionNode { Name: "IF", Arguments.Count: 2 or 3 } choice)
        {
            var condition = Scalar(choice.Arguments[0]); if (condition.Kind == CellValueKind.Error) Fail(condition.ErrorCode, "SUBTOTAL 引用条件包含错误。");
            return SubtotalReference(Boolean(condition) ? choice.Arguments[1] : choice.Arguments.Count == 3 ? choice.Arguments[2] : new MissingNode());
        }
        if (node is FunctionNode { Name: "INDEX", Arguments.Count: 2 or 3 } index && index.Arguments[0] is ReferenceNode source)
        {
            double Coordinate(int i) { var value = Scalar(index.Arguments[i]); if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "SUBTOTAL INDEX 包含错误。"); return Math.Truncate(Number(value)); }
            var row = Coordinate(1); var column = index.Arguments.Count == 3 ? Coordinate(2) : 1; var (rows, columns) = ConditionalRanges.Shape(source);
            if (rows == 1 && index.Arguments.Count == 2) { column = row; row = 1; }
            if (row < 0 || column < 0 || row > rows || column > columns) Fail(CellErrorCode.Reference, "SUBTOTAL INDEX 超出区域。");
            return SubtotalReference(source with { Row = row == 0 ? source.Row : source.Row + (int)row - 1, EndRow = row == 0 ? source.EndRow : source.Row + (int)row - 1,
                Column = column == 0 ? source.Column : source.Column + (int)column - 1, EndColumn = column == 0 ? source.EndColumn : source.Column + (int)column - 1, IsRange = true, WholeColumns = row == 0 && source.WholeColumns });
        }
        if (node is ValueNode { Value.Kind: CellValueKind.Error } error) Fail(error.Value.ErrorCode, "SUBTOTAL 引用错误。");
        throw new FormulaValueException(CellErrorCode.Value, "SUBTOTAL 需要单元格或区域引用。");
    }
    private bool ExcludedByFilter(WorksheetSnapshot sheet, int row)
    {
        if (_subtotalFiltered.TryGetValue((sheet.Id, row), out var excluded)) return excluded;
        foreach (var filter in sheet.Data.AllFilters)
        {
            WorksheetFilterEngine.RequireColorScope(sheet, filter);
            if (row <= filter.Range.Top || row > filter.Range.Bottom) continue;
            foreach (var pair in filter.Columns)
            {
                check(); var value = Read(sheet, row, pair.Key);
                if (!_filterPredicates.TryGetValue(pair.Value, out var predicate)) _filterPredicates[pair.Value] = predicate = WorksheetFilterEngine.Compile(pair.Value);
                if (!predicate(value, TablePresentation.NumberStyle(sheet, snapshot.Styles, row, pair.Key), snapshot.DateSystem)) return _subtotalFiltered[(sheet.Id, row)] = true;
            }
        }
        return _subtotalFiltered[(sheet.Id, row)] = false;
    }
    private bool NestedSubtotal(WorksheetSnapshot sheet, CellPosition p)
    {
        var cell = sheet.GetCell(p.Row, p.Column); if (cell?.Content.Kind != CellContentKind.Formula) return false;
        var address = new CellAddress(sheet.Id, p.Row, p.Column);
        if (_nestedSubtotals.TryGetValue(address, out var nested)) return nested;
        return _nestedSubtotals[address] = SubtotalFunctions.Contains(FormulaBinding.Bind(FormulaSyntax.Parse(cell.Content.FormulaText), snapshot, sheet.Id, p));
    }
    private CellValue Subtotal(FunctionNode function)
    {
        var codeValue = function.Arguments[0] is MissingNode ? CellValue.Blank : Scalar(function.Arguments[0]); if (codeValue.Kind == CellValueKind.Error) return codeValue;
        var code = Math.Truncate(Number(codeValue)); if (code is not (>= 1 and <= 11 or >= 101 and <= 111)) return Error(CellErrorCode.Value);
        var excludeHidden = code >= 101; var operation = (int)code % 100;
        double sum = 0, product = 1, minimum = double.PositiveInfinity, maximum = double.NegativeInfinity, mean = 0, m2 = 0;
        var count = 0; var nonempty = 0;
        foreach (var arg in function.Arguments.Skip(1))
        {
            var range = SubtotalReference(arg); var sheet = range.Sheet;
            if (!_positions.TryGetValue(sheet.Id, out var positions)) _positions[sheet.Id] = positions = sheet.Cells.Keys.OrderBy(p => p.Row).ThenBy(p => p.Column).ToArray();
            foreach (var p in positions)
            {
                check(); if (p.Row > range.Reference.EndRow) break;
                if (!range.Reference.Contains(p.Row, p.Column) || excludeHidden && sheet.Rows.GetValueOrDefault(p.Row)?.Hidden == true ||
                    NestedSubtotal(sheet, p) || ExcludedByFilter(sheet, p.Row)) continue;
                var value = Read(sheet, p.Row, p.Column);
                if (value.Kind != CellValueKind.Blank) nonempty++;
                if (operation == 3) continue;
                if (value.Kind == CellValueKind.Error) { if (operation == 2) continue; return value; }
                if (value.Kind != CellValueKind.Number) continue;
                var number = value.NumberValue; count++; sum += number; product *= number; minimum = Math.Min(minimum, number); maximum = Math.Max(maximum, number);
                var delta = number - mean; mean += delta / count; m2 += delta * (number - mean);
            }
        }
        if (operation is 7 or 10 && count < 2 || operation is 1 or 8 or 11 && count == 0) return Error(CellErrorCode.DivisionByZero);
        return CellValue.FromNumber(operation switch
        {
            1 => sum / count, 2 => count, 3 => nonempty, 4 => count == 0 ? 0 : maximum, 5 => count == 0 ? 0 : minimum,
            6 => count == 0 ? 0 : product, 7 => Math.Sqrt(Math.Max(0, m2 / (count - 1))), 8 => Math.Sqrt(Math.Max(0, m2 / count)),
            9 => sum, 10 => m2 / (count - 1), _ => m2 / count
        });
    }
}
