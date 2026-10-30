using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private Range ConditionalRange(FormulaNode node)
    {
        if (node is ReferenceNode reference)
            return new(snapshot.Sheets.First(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase)), reference);
        if (node is ValueNode { Value.Kind: CellValueKind.Error } error) Fail(error.Value.ErrorCode, "统计区域引用无效。");
        throw new FormulaValueException(CellErrorCode.Value, "条件统计需要矩形单元格区域。");
    }

    private CellValue Conditional(FunctionNode function)
    {
        var args = function.Arguments; var name = function.Name;
        var count = name is "COUNTIF" or "COUNTIFS" or "COUNTBLANK";
        var paired = name.EndsWith("IFS", StringComparison.Ordinal);
        if (paired && args.Count % 2 != (name == "COUNTIFS" ? 0 : 1)) return Error(CellErrorCode.Value);
        var ranges = new List<Range>(); var criteria = new List<FormulaCriterion>(); Range? output = null;
        FormulaCriterion Criterion(FormulaNode node)
        {
            CellValue value;
            try
            {
                value = node is ReferenceNode r && ConditionalRanges.Shape(r) == (1, 1)
                    ? Read(ConditionalRange(r).Sheet, r.Row, r.Column) : Scalar(node);
            }
            catch (FormulaValueException ex) { value = Error(ex.Code); }
            return FormulaCriterion.Create(value, snapshot.DateSystem);
        }
        if (name == "COUNTBLANK") ranges.Add(ConditionalRange(args[0]));
        else if (!paired)
        {
            ranges.Add(ConditionalRange(args[0])); criteria.Add(Criterion(args[1]));
            if (!count)
            {
                output = args.Count == 2 ? ranges[0] : ConditionalRange(args[2]);
                output = output with { Reference = ConditionalRanges.Project(ranges[0].Reference, output.Reference) };
            }
        }
        else
        {
            if (!count) output = ConditionalRange(args[0]);
            for (var i = count ? 0 : 1; i < args.Count; i += 2)
            { ranges.Add(ConditionalRange(args[i])); criteria.Add(Criterion(args[i + 1])); }
        }
        var (rows, columns) = ConditionalRanges.Shape(ranges[0].Reference);
        if (paired && (ranges.Any(r => ConditionalRanges.Shape(r.Reference) != (rows, columns)) ||
            output is not null && ConditionalRanges.Shape(output.Reference) != (rows, columns))) return Error(CellErrorCode.Value);

        // Only offsets with stored input need visiting. All other rows have the same blank tuple.
        // The union includes the projected result range so sums for blank criteria remain correct.
        var offsets = new HashSet<long>();
        foreach (var range in output is null ? ranges : ranges.Append(output))
        {
            if (!_positions.TryGetValue(range.Sheet.Id, out var positions))
                _positions[range.Sheet.Id] = positions = range.Sheet.Cells.Keys.OrderBy(p => p.Row).ThenBy(p => p.Column).ToArray();
            foreach (var p in positions)
            {
                check();
                if (p.Row > range.Reference.EndRow) break;
                if (!range.Reference.Contains(p.Row, p.Column)) continue;
                offsets.Add((long)(p.Row - range.Reference.Row) * columns + p.Column - range.Reference.Column);
                if (offsets.Count > WorkbookLimits.MaxCells) Fail(CellErrorCode.Value, "条件统计的稀疏位置并集超过 1,000,000 项。");
            }
        }
        static bool Blank(CellValue value) => value.Kind == CellValueKind.Blank || value is { Kind: CellValueKind.Text, TextValue.Length: 0 };
        var total = count && (name == "COUNTBLANK" || criteria.All(c => c.Matches(CellValue.Blank))) ? (long)rows * columns - offsets.Count : 0L;
        double sum = 0; long numeric = 0;
        CellValue At(Range range, int row, int column) => row + range.Reference.Row >= ConditionalRanges.ExcelRows || column + range.Reference.Column >= ConditionalRanges.ExcelColumns
            ? CellValue.Blank : Read(range.Sheet, range.Reference.Row + row, range.Reference.Column + column);
        foreach (var offset in offsets.Order())
        {
            check(); var row = (int)(offset / columns); var column = (int)(offset % columns);
            if (name == "COUNTBLANK") { if (Blank(At(ranges[0], row, column))) total++; continue; }
            var matches = true;
            for (var i = 0; i < ranges.Count; i++)
                if (!criteria[i].Matches(At(ranges[i], row, column))) { matches = false; break; }
            if (!matches) continue;
            if (count) { total++; continue; }
            var value = At(output!, row, column);
            if (value.Kind == CellValueKind.Error) return value;
            if (value.Kind == CellValueKind.Number) { sum += value.NumberValue; numeric++; }
        }
        if (count) return CellValue.FromNumber(total);
        return name.StartsWith("AVERAGE", StringComparison.Ordinal)
            ? numeric == 0 ? Error(CellErrorCode.DivisionByZero) : CellValue.FromNumber(sum / numeric)
            : CellValue.FromNumber(sum);
    }
}
