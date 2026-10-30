using System.Globalization;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

/// <summary>Checks the candidate immutable transaction, evaluating only required standard-formula dependencies.</summary>
public sealed class WorkbookValidation : IWorkbookValidation
{
    public IReadOnlyList<ValidationFailure> Check(WorkbookSnapshot snapshot, IReadOnlyList<CellAddress> addresses)
    {
        var evaluation = new Evaluation(snapshot); var failures = new List<ValidationFailure>();
        foreach (var address in addresses)
        {
            var sheet = snapshot.GetSheet(address.SheetId); var position = new CellPosition(address.Row, address.Column);
            foreach (var rule in sheet.Data.Validations.Where(v => v.Contains(position) && v.ShowError && v.Kind != ValidationKind.None))
            {
                try
                {
                    if (Valid(evaluation, address, position, rule)) continue;
                    failures.Add(new(address, rule.Alert, rule.ErrorTitle ?? "数据验证", rule.ErrorMessage ?? "输入值不符合此单元格的验证规则。"));
                }
                catch (Exception ex) when (ex is FormatException or NotSupportedException or InvalidOperationException)
                { failures.Add(new(address, ValidationAlert.Stop, "无法确认验证结果", ex.Message)); }
            }
        }
        return failures;
    }

    public IReadOnlyList<CellValue> ListValues(WorkbookSnapshot snapshot, SheetId sheet, CellPosition position, WorksheetValidation rule) =>
        new Evaluation(snapshot).List(sheet, position, rule, out _);

    private static bool Valid(Evaluation evaluation, CellAddress address, CellPosition position, WorksheetValidation rule)
    {
        var value = evaluation.Cell(address);
        if (rule.AllowBlank && value.Kind == CellValueKind.Blank) return true;
        if (value.Kind == CellValueKind.Error) return false;
        if (rule.Kind == ValidationKind.List)
        {
            var choices = evaluation.List(address.SheetId, position, rule, out var inline);
            // Excel treats a blank in a named source differently from a direct range source.
            if (rule.AllowBlank && FormulaSyntax.Parse(rule.Formula1!) is NameNode && choices.Any(v => v.Kind == CellValueKind.Blank)) return true;
            if (value.Kind == CellValueKind.Blank) return false;
            return choices.Any(v => v.Kind is not (CellValueKind.Error or CellValueKind.Blank) &&
                (inline ? string.Equals(CellInput.DisplayText(v), CellInput.DisplayText(value), StringComparison.Ordinal) :
                    v.Kind == CellValueKind.Text && value.Kind == CellValueKind.Text ? string.Equals(v.TextValue, value.TextValue, StringComparison.OrdinalIgnoreCase) : v == value));
        }
        CellValue Formula(string source) => evaluation.Scalar(source, address.SheetId, rule.Anchor, position);
        if (rule.AllowBlank && new[] { rule.Formula1, rule.Formula2 }.OfType<string>().Any(f => evaluation.HasBlankReference(f, address.SheetId, rule.Anchor, position)))
            return true;
        if (rule.Kind == ValidationKind.Custom)
        {
            var result = Formula(rule.Formula1!);
            if (result.Kind == CellValueKind.Error) throw new NotSupportedException("自定义验证公式计算出错误，不能确认输入是否有效。");
            return result.Kind == CellValueKind.Boolean ? result.BooleanValue : result.Kind == CellValueKind.Number && result.NumberValue != 0;
        }
        double number;
        if (rule.Kind == ValidationKind.TextLength)
        {
            number = evaluation.Length(value, address.SheetId);
        }
        else
        {
            if (value.Kind == CellValueKind.Blank) number = 0;
            else if (value.Kind == CellValueKind.Number) number = value.NumberValue;
            else if (value.Kind == CellValueKind.Boolean) number = double.PositiveInfinity;
            else return false;
            if (rule.Kind == ValidationKind.Whole && number != Math.Truncate(number)) return false;
        }
        var first = Formula(rule.Formula1!);
        if (first.Kind == CellValueKind.Error) throw new NotSupportedException("验证条件公式计算出错误。");
        if (rule.AllowBlank && first.Kind == CellValueKind.Blank) return true;
        if (first.Kind is not (CellValueKind.Number or CellValueKind.Blank)) return false;
        var lower = first.Kind == CellValueKind.Blank ? 0 : first.NumberValue;
        var upper = 0d;
        if (rule.Comparison is ValidationOperator.Between or ValidationOperator.NotBetween)
        {
            var second = Formula(rule.Formula2!);
            if (second.Kind == CellValueKind.Error) throw new NotSupportedException("验证条件公式计算出错误。");
            if (rule.AllowBlank && second.Kind == CellValueKind.Blank) return true;
            if (second.Kind is not (CellValueKind.Number or CellValueKind.Blank)) return false;
            upper = second.Kind == CellValueKind.Blank ? 0 : second.NumberValue;
        }
        return rule.Comparison switch
        {
            ValidationOperator.Between => number >= lower && number <= upper,
            ValidationOperator.NotBetween => number < lower || number > upper,
            ValidationOperator.Equal => number == lower,
            ValidationOperator.NotEqual => number != lower,
            ValidationOperator.LessThan => number < lower,
            ValidationOperator.LessThanOrEqual => number <= lower,
            ValidationOperator.GreaterThan => number > lower,
            ValidationOperator.GreaterThanOrEqual => number >= lower,
            _ => false
        };
    }

    private sealed class Evaluation(WorkbookSnapshot snapshot)
    {
        private readonly Dictionary<CellAddress, CellCalculation> _results = [];
        private readonly HashSet<CellAddress> _active = [];
        private int _work;
        private void Check() { if (++_work > 2_000_000) throw new NotSupportedException("验证计算超过 2,000,000 步，请缩小编辑范围。"); }
        private FormulaNode Bind(string formula, SheetId sheet, CellPosition anchor, CellPosition position, bool list = false)
        {
            Check();
            var expression = FormulaSyntax.Parse(FormulaCopy.Translate(formula, anchor, position));
            // A list range is consumed as a range, rather than as a scalar cell formula.
            var bound = FormulaBinding.Bind(list ? new FunctionNode("SUM", [expression]) : expression, snapshot, sheet, position);
            if (list) bound = ((FunctionNode)bound).Arguments[0];
            if (FormulaSyntax.Walk(bound).OfType<FunctionNode>().Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) ||
                f.Name is "NOW" or "TODAY" or "RAND" or "RANDBETWEEN"))
                throw new NotSupportedException("验证暂不执行脚本、未知或易变函数。");
            return bound;
        }
        public CellValue Cell(CellAddress address)
        {
            Check();
            if (_results.TryGetValue(address, out var previous)) return previous.Value ?? CellValue.Blank;
            var cell = snapshot.GetSheet(address.SheetId).GetCell(address.Row, address.Column);
            if (cell?.Content.Kind != CellContentKind.Formula) return cell?.Value ?? CellValue.Blank;
            if (_active.Count >= 128 || !_active.Add(address)) throw new NotSupportedException("验证依赖存在循环或嵌套超过 128 层。");
            try
            {
                var position = new CellPosition(address.Row, address.Column);
                var value = Evaluate(Bind(cell.Content.FormulaText, address.SheetId, position, position), address.SheetId, cellResult: true);
                _results[address] = new(value, value.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh, 0, 0, 0);
                return value;
            }
            finally { _active.Remove(address); }
        }
        private CellValue Evaluate(FormulaNode expression, SheetId sheet, bool cellResult = false)
        {
            foreach (var reference in FormulaSyntax.Walk(expression).OfType<ReferenceNode>())
            {
                var owner = snapshot.Sheets.Single(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
                if (!reference.IsRange && !reference.WholeColumns) { _ = Cell(new(owner.Id, reference.Row, reference.Column)); continue; }
                foreach (var p in owner.Cells)
                {
                    Check();
                    if (p.Value.Content.Kind == CellContentKind.Formula && reference.Contains(p.Key.Row, p.Key.Column)) _ = Cell(new(owner.Id, p.Key.Row, p.Key.Column));
                }
            }
            var evaluator = new TypedEvaluator(snapshot, _results, null, new Dictionary<string, ScriptFunctionDescriptor>(), Check) { CurrentSheet = sheet };
            try { return cellResult ? evaluator.CellResult(expression) : evaluator.Scalar(expression); }
            catch (FormulaValueException ex) { return CellValue.FromError(ex.Code); }
        }
        public CellValue Scalar(string formula, SheetId sheet, CellPosition anchor, CellPosition position) => Evaluate(Bind(formula, sheet, anchor, position), sheet);
        public bool HasBlankReference(string formula, SheetId sheet, CellPosition anchor, CellPosition position) =>
            FormulaSyntax.Walk(Bind(formula, sheet, anchor, position)).OfType<ReferenceNode>().Any(r => !r.IsRange && !r.WholeColumns &&
                Cell(new(snapshot.Sheets.Single(s => s.Name.Equals(r.Sheet, StringComparison.OrdinalIgnoreCase)).Id, r.Row, r.Column)).Kind == CellValueKind.Blank);
        public double Length(CellValue value, SheetId sheet) => Evaluate(new FunctionNode("LEN", [new ValueNode(value)]), sheet).NumberValue;
        public IReadOnlyList<CellValue> List(SheetId sheet, CellPosition position, WorksheetValidation rule, out bool inline)
        {
            var bound = Bind(rule.Formula1!, sheet, rule.Anchor, position, true);
            inline = bound is ValueNode { Value.Kind: CellValueKind.Text };
            if (bound is ValueNode { Value.Kind: CellValueKind.Text } literal) return literal.Value.TextValue.Split(',').Select(CellValue.FromText).ToArray();
            if (bound is not ReferenceNode reference || reference.Row != reference.EndRow && reference.Column != reference.EndColumn ||
                (long)(reference.EndRow - reference.Row + 1) * (reference.EndColumn - reference.Column + 1) > 10000)
                throw new NotSupportedException("列表来源须为不超过 10,000 格的单行或单列区域。");
            var owner = snapshot.Sheets.Single(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
            var result = new List<CellValue>();
            for (var r = reference.Row; r <= reference.EndRow; r++)
            for (var c = reference.Column; c <= reference.EndColumn; c++) result.Add(Cell(new(owner.Id, r, c)));
            return result;
        }
    }
}
