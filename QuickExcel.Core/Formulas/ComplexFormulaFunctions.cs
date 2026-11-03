using System.Globalization;
using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Formulas;

/// <summary>Reference functions whose complete target is known before dependency planning. Never reads caches or external files.</summary>
public static class ComplexFormulaFunctions
{
    public static bool IsStatistic(string name) => name is "VAR.S" or "VAR.P" or "STDEV.S" or "STDEV.P" or "MEDIAN";
    public static bool IsLocalName(string name) => name.Length is > 0 and <= 255 &&
        (char.IsLetter(name[0]) || name[0] is '_' or '\\') && name.All(c => char.IsLetterOrDigit(c) || c is '_' or '.' or '\\') &&
        !name.Equals("R", StringComparison.OrdinalIgnoreCase) && !name.Equals("C", StringComparison.OrdinalIgnoreCase) &&
        !FormulaSyntax.IsCellName(name) && !Regex.IsMatch(name, @"^R[0-9]+C[0-9]+$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static FormulaNode Constant(FormulaNode node)
    {
        for (var depth = 0; node is LocalValueNode local; depth++)
        {
            if (depth >= 128) throw new FormatException("局部名称的引用链超过 128 层。");
            node = local.Definition;
        }
        return node;
    }
    private static ValueNode Error(CellErrorCode code) => new(CellValue.FromError(code));
    public static FormulaNode BindReference(FunctionNode function, WorkbookSnapshot book, SheetId sheetId, CellPosition? position)
    {
        var args = function.Arguments;
        if (!StandardFunctions.AcceptsArguments(function.Name, args.Count)) return Error(CellErrorCode.Value);
        if (function.Name == "OFFSET")
        {
            if (args[0] is ValueNode { Value.Kind: CellValueKind.Error } error) return error;
            if (args[0] is not ReferenceNode source) throw new FormatException("OFFSET 基准须为可静态确定的本地矩形引用。");
            var size = ConditionalRanges.Shape(source); var values = new double[] { 0, 0, size.Rows, size.Columns };
            for (var i = 1; i < args.Count; i++)
            {
                var argument = Constant(args[i]);
                if (argument is MissingNode) { values[i - 1] = i >= 3 ? 1 : 0; continue; }
                if (argument is ValueNode { Value.Kind: CellValueKind.Error } e) return e;
                if (!LookupModes.TryConstant(argument, out var number))
                {
                    if (argument is ValueNode) return Error(CellErrorCode.Value);
                    throw new FormatException("OFFSET 偏移和尺寸目前须为常量；动态依赖目标保留来源公式和缓存。");
                }
                values[i - 1] = Math.Truncate(number);
            }
            var row = source.Row + values[0]; var column = source.Column + values[1];
            var endRow = row + values[2] - Math.Sign(values[2]); var endColumn = column + values[3] - Math.Sign(values[3]);
            if (values[2] == 0 || values[3] == 0 || Math.Min(row, endRow) < 0 || Math.Min(column, endColumn) < 0 ||
                Math.Max(row, endRow) >= ConditionalRanges.ExcelRows || Math.Max(column, endColumn) >= ConditionalRanges.ExcelColumns)
                return Error(CellErrorCode.Reference);
            if (Math.Max(row, endRow) >= WorkbookLimits.MaxRows || Math.Max(column, endColumn) >= WorkbookLimits.MaxColumns)
                throw new FormatException("OFFSET 目标超过产品坐标上限；保留来源公式和缓存。");
            return new ReferenceNode(source.Sheet, (int)Math.Min(row, endRow), (int)Math.Min(column, endColumn),
                (int)Math.Max(row, endRow), (int)Math.Max(column, endColumn), IsRange: Math.Abs(values[2]) != 1 || Math.Abs(values[3]) != 1, SourceExpression: function);
        }
        var textNode = Constant(args[0]);
        if (textNode is ValueNode { Value.Kind: CellValueKind.Error } textError) return textError;
        if (textNode is not ValueNode { Value.Kind: CellValueKind.Text } text)
        {
            if (textNode is ValueNode or MissingNode) return Error(CellErrorCode.Reference);
            throw new FormatException("INDIRECT 目前须为常量地址文本；动态依赖目标保留来源公式和缓存。");
        }
        var a1 = args.Count != 2 || Constant(args[1]) is not MissingNode;
        if (args.Count == 2 && Constant(args[1]) is not MissingNode)
        {
            var mode = Constant(args[1]);
            if (mode is ValueNode { Value.Kind: CellValueKind.Error } modeError) return modeError;
            if (!LookupModes.TryConstant(mode, out var number))
            {
                if (mode is ValueNode) return Error(CellErrorCode.Value);
                throw new FormatException("INDIRECT 地址模式须为常量。");
            }
            a1 = number != 0;
        }
        var address = text.Value.TextValue;
        if (book.DefinedNames.Any(n => n.Name.Equals(address, StringComparison.OrdinalIgnoreCase)))
            throw new FormatException("INDIRECT 名称地址尚未建模；保留来源公式和缓存。");
        if (a1 && (address.Contains('[') || address.Contains(']')) || address.Contains("://", StringComparison.Ordinal))
            throw new FormatException("INDIRECT 外部/结构化地址尚未建模；不会访问外部内容。");
        if (!a1)
        {
            var bang = address.LastIndexOf('!'); var prefix = bang < 0 ? "" : address[..(bang + 1)]; var cells = address[(bang + 1)..].Split(':');
            if (prefix.Contains('[') || prefix.Contains(']')) throw new FormatException("INDIRECT 外部地址尚未建模。");
            if (cells.Length > 2) return Error(CellErrorCode.Reference);
            var converted = new List<string>();
            foreach (var cell in cells)
            {
                var match = Regex.Match(cell, @"^R(?<row>\[[+-]?\d+\]|\d*)C(?<col>\[[+-]?\d+\]|\d*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!match.Success)
                {
                    if (Regex.IsMatch(cell, @"^[RC](\[[+-]?\d+\]|\d*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                        throw new FormatException("INDIRECT R1C1 整行/整列地址尚未建模。");
                    return Error(CellErrorCode.Reference);
                }
                long Axis(string value, int? origin)
                {
                    if (value == "" || value.StartsWith('['))
                    {
                        if (origin is null) throw new FormatException("相对 R1C1 地址需要明确公式位置。");
                        return origin.Value + (value == "" ? 0 : long.Parse(value[1..^1], CultureInfo.InvariantCulture));
                    }
                    return long.Parse(value, CultureInfo.InvariantCulture) - 1;
                }
                long row, column;
                try { row = Axis(match.Groups["row"].Value, position?.Row); column = Axis(match.Groups["col"].Value, position?.Column); }
                catch (OverflowException) { return Error(CellErrorCode.Reference); }
                if (row < 0 || row >= ConditionalRanges.ExcelRows || column < 0 || column >= ConditionalRanges.ExcelColumns) return Error(CellErrorCode.Reference);
                converted.Add(FormulaSyntax.ColumnName((int)column) + (row + 1).ToString(CultureInfo.InvariantCulture));
            }
            address = prefix + string.Join(':', converted);
        }
        FormulaNode parsed;
        try { parsed = FormulaSyntax.Parse("=" + address); }
        catch (FormatException e) when (e.Message.Contains("坐标范围", StringComparison.Ordinal))
        { throw new FormatException("INDIRECT 目标超过产品坐标上限；保留来源公式和缓存。", e); }
        catch (FormatException) when (string.IsNullOrWhiteSpace(address)) { return Error(CellErrorCode.Reference); }
        if (parsed is NameNode name && FormulaBinding.ResolveName(book, sheetId, name) is not null)
            throw new FormatException("INDIRECT 名称地址尚未建模；保留来源公式和缓存。");
        if (parsed is not ReferenceNode target) return Error(CellErrorCode.Reference);
        var sheet = target.Sheet is null ? book.GetSheet(sheetId) : book.Sheets.FirstOrDefault(s => s.Name.Equals(target.Sheet, StringComparison.OrdinalIgnoreCase));
        return sheet is null ? Error(CellErrorCode.Reference) : target with { Sheet = sheet.Name, SourceExpression = function };
    }
}
