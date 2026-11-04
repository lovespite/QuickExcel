using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Formulas;

public abstract record FormulaNode;
public sealed record MissingNode : FormulaNode;
public sealed record ValueNode(CellValue Value) : FormulaNode;
public sealed record NameNode(string Name, string? Sheet = null) : FormulaNode;
public sealed record UnaryNode(string Operator, FormulaNode Operand) : FormulaNode;
public sealed record BinaryNode(string Operator, FormulaNode Left, FormulaNode Right) : FormulaNode;
public sealed record FunctionNode(string Name, IReadOnlyList<FormulaNode> Arguments) : FormulaNode;
// Bound-only lexical nodes. Local references never expand their expression at each use.
public sealed record LocalValueNode(int Id, string Name, FormulaNode Definition, bool ArrayResult, int EvaluationDepth) : FormulaNode;
public sealed record LetBinding(int Id, string Name, FormulaNode Value);
public sealed record LetNode(IReadOnlyList<LetBinding> Bindings, FormulaNode Body) : FormulaNode;
public sealed record ReferenceNode(string? Sheet, int Row, int Column, int EndRow, int EndColumn,
    bool ColumnAbsolute = false, bool RowAbsolute = false, bool EndColumnAbsolute = false,
    bool EndRowAbsolute = false, bool IsRange = false, bool WholeColumns = false,
    FormulaNode? SourceExpression = null) : FormulaNode
{
    public bool Contains(int row, int column) => row >= Row && row <= EndRow && column >= Column && column <= EndColumn;
}

/// <summary>Strict, bounded A1 syntax shared by binding and structural commands. No IO or runtime dependency.</summary>
public static class FormulaSyntax
{
    public static FormulaNode Parse(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        if (formula.Length > 8192) throw new FormatException("公式最多支持 8192 个字符。");
        return new Parser(formula).Parse();
    }

    public static IEnumerable<FormulaNode> Walk(FormulaNode root)
    {
        var stack = new Stack<(FormulaNode Node, HashSet<string>? Locals)>();
        stack.Push((root, null));
        while (stack.TryPop(out var item))
        {
            var (node, locals) = item;
            if (node is NameNode { Sheet: null } local && locals?.Contains(local.Name) == true) continue;
            yield return node;
            switch (node)
            {
                case UnaryNode unary: stack.Push((unary.Operand, locals)); break;
                case BinaryNode binary: stack.Push((binary.Right, locals)); stack.Push((binary.Left, locals)); break;
                case LetNode let:
                    stack.Push((let.Body, locals));
                    for (var i = let.Bindings.Count - 1; i >= 0; i--) stack.Push((let.Bindings[i].Value, locals));
                    break;
                case FunctionNode { Name: "LET" } let when StandardFunctions.AcceptsArguments("LET", let.Arguments.Count):
                    var children = new List<(FormulaNode, HashSet<string>?)>(); var scope = locals;
                    for (var i = 0; i < let.Arguments.Count - 1; i += 2)
                    {
                        children.Add((let.Arguments[i + 1], scope));
                        scope = scope is null ? new(StringComparer.OrdinalIgnoreCase) : new(scope, StringComparer.OrdinalIgnoreCase);
                        if (let.Arguments[i] is NameNode { Sheet: null } name) scope.Add(name.Name);
                    }
                    children.Add((let.Arguments[^1], scope));
                    for (var i = children.Count - 1; i >= 0; i--) stack.Push(children[i]);
                    break;
                case FunctionNode function:
                    for (var i = function.Arguments.Count - 1; i >= 0; i--) stack.Push((function.Arguments[i], locals));
                    break;
            }
        }
    }

    public static FormulaNode Map(FormulaNode node, Func<FormulaNode, FormulaNode> transform) => MapScoped(node, transform, null);
    private static FormulaNode MapScoped(FormulaNode node, Func<FormulaNode, FormulaNode> transform, HashSet<string>? locals)
    {
        if (node is NameNode { Sheet: null } local && locals?.Contains(local.Name) == true) return node;
        if (node is FunctionNode { Name: "LET" } let && StandardFunctions.AcceptsArguments("LET", let.Arguments.Count))
        {
            var scope = locals is null ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : new(locals, StringComparer.OrdinalIgnoreCase);
            var args = let.Arguments.ToArray();
            for (var i = 0; i < args.Length - 1; i += 2)
            {
                args[i + 1] = MapScoped(args[i + 1], transform, scope);
                if (args[i] is NameNode { Sheet: null } name) scope.Add(name.Name);
            }
            args[^1] = MapScoped(args[^1], transform, scope);
            return transform(let with { Arguments = args });
        }
        FormulaNode Child(FormulaNode child) => MapScoped(child, transform, locals);
        var transformed = transform(node switch
        {
            UnaryNode unary => unary with { Operand = Child(unary.Operand) },
            BinaryNode binary => binary with { Left = Child(binary.Left), Right = Child(binary.Right) },
            LetNode bound => bound with { Bindings = bound.Bindings.Select(b => b with { Value = Child(b.Value) }).ToArray(), Body = Child(bound.Body) },
            FunctionNode function => function with { Arguments = function.Arguments.Select(Child).ToArray() },
            _ => node
        });
        if (node is NameNode && transformed is NameNode { Sheet: null } renamed && locals?.Contains(renamed.Name) == true)
            throw new NotSupportedException("名称转换会被 LET 局部变量遮蔽，操作未应用；请使用不冲突的名称。");
        return transformed;
    }

    public static string ColumnName(int column)
    {
        var name = "";
        for (var n = column + 1; n > 0; n = (n - 1) / 26) name = (char)('A' + (n - 1) % 26) + name;
        return name;
    }

    public static string SheetPrefix(string? sheet) => sheet is null ? "" : "'" + sheet.Replace("'", "''") + "'!";

    /// <summary>Print an expanded formula with a bounded allocation; false means it cannot be retained in a portable file.</summary>
    public static bool TryPrintPortable(FormulaNode node, out string formula)
    {
        const int limit = 8192;
        var output = new StringBuilder("="); var nodes = 0;
        bool Append(string text)
        {
            if (text.Length > limit - output.Length) return false;
            output.Append(text); return true;
        }
        bool Write(FormulaNode part, int depth)
        {
            if (depth > 128 || ++nodes > 2048) return false;
            switch (part)
            {
                case MissingNode: return true;
                case ValueNode { Value.Kind: CellValueKind.Text } value:
                    if (!Append("\"")) return false;
                    foreach (var ch in value.Value.TextValue)
                        if (!Append(ch == '"' ? "\"\"" : ch.ToString())) return false;
                    return Append("\"");
                case ValueNode or NameNode or ReferenceNode or StructuredReferenceNode or LocalValueNode: return Append(Print(part));
                case LetNode let: return Write(LetFunction(let), depth + 1);
                case UnaryNode unary:
                    return Append(unary.Operator is "%" or "#" ? "(" : unary.Operator + "(") && Write(unary.Operand, depth + 1)
                        && Append(unary.Operator is "%" or "#" ? ")" + unary.Operator : ")");
                case BinaryNode binary:
                    return Append("(") && Write(binary.Left, depth + 1) && Append(binary.Operator) && Write(binary.Right, depth + 1) && Append(")");
                case FunctionNode function:
                    if (!Append(function.Name) || !Append("(")) return false;
                    for (var i = 0; i < function.Arguments.Count; i++)
                        if (i > 0 && !Append(",") || !Write(function.Arguments[i], depth + 1)) return false;
                    return Append(")");
                default: return false;
            }
        }
        formula = "";
        if (!Write(node, 0)) return false;
        var result = output.ToString();
        if (ExcelFormulaNames.Encode(result).Length > limit) return false;
        // Printing inserts parentheses and qualified sheet names. Confirm our own reader can reopen the result.
        try { _ = Parse(result); } catch (FormatException) { return false; }
        formula = result; return true;
    }

    public static string Print(FormulaNode node) => node switch
    {
        MissingNode => "",
        ValueNode v => v.Value.Kind switch
        {
            CellValueKind.Text => "\"" + v.Value.TextValue.Replace("\"", "\"\"") + "\"",
            CellValueKind.Number => v.Value.NumberValue.ToString("R", CultureInfo.InvariantCulture),
            CellValueKind.Boolean => v.Value.BooleanValue ? "TRUE" : "FALSE",
            CellValueKind.Error => CellInput.DisplayText(v.Value),
            _ => "\"\""
        },
        NameNode n => SheetPrefix(n.Sheet) + n.Name,
        ReferenceNode r => PrintReference(r),
        StructuredReferenceNode r => StructuredReferences.Print(r),
        UnaryNode u => u.Operator is "%" or "#" ? "(" + Print(u.Operand) + ")" + u.Operator : u.Operator + "(" + Print(u.Operand) + ")",
        BinaryNode b => "(" + Print(b.Left) + b.Operator + Print(b.Right) + ")",
        FunctionNode f => f.Name + "(" + string.Join(",", f.Arguments.Select(Print)) + ")",
        LocalValueNode local => local.Name,
        LetNode let => Print(LetFunction(let)),
        _ => throw new NotSupportedException()
    };

    private static string PrintReference(ReferenceNode r)
    {
        if (r.SourceExpression is not null) return Print(r.SourceExpression);
        string End(int row, int col, bool ca, bool ra) => (ca ? "$" : "") + ColumnName(col)
            + (r.WholeColumns ? "" : (ra ? "$" : "") + (row + 1).ToString(CultureInfo.InvariantCulture));
        return SheetPrefix(r.Sheet) + End(r.Row, r.Column, r.ColumnAbsolute, r.RowAbsolute)
            + (r.IsRange || r.WholeColumns ? ":" + End(r.EndRow, r.EndColumn, r.EndColumnAbsolute, r.EndRowAbsolute) : "");
    }

    private static FunctionNode LetFunction(LetNode let) => new("LET", let.Bindings
        .SelectMany(b => new FormulaNode[] { new NameNode(b.Name), b.Value }).Append(let.Body).ToArray());

    private static readonly Regex CellPattern = new(@"^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$", RegexOptions.CultureInvariant);
    private static readonly Regex ColumnPattern = new(@"^(\$?)([A-Za-z]{1,3})$", RegexOptions.CultureInvariant);
    public static bool IsCellName(string value) => CellPattern.IsMatch(value);

    private readonly record struct Token(string Kind, string Text);
    private sealed class Parser
    {
        private readonly List<Token> _tokens = [];
        private int _index;
        private int _nodes;
        private Token Current => _tokens[_index];
        private Token Take() => _tokens[_index++];
        public Parser(string input)
        {
            var pos = input.StartsWith('=') ? 1 : 0;
            while (pos < input.Length)
            {
                var ch = input[pos];
                if (char.IsWhiteSpace(ch)) { pos++; continue; }
                if (ch == '[') { _tokens.Add(new("table", StructuredReferences.Token(input, ref pos))); continue; }
                if (ch is '"' or '\'')
                {
                    var quote = ch; var text = new StringBuilder(); pos++; var closed = false;
                    while (pos < input.Length)
                    {
                        ch = input[pos++];
                        if (ch != quote) { text.Append(ch); continue; }
                        if (pos < input.Length && input[pos] == quote) { text.Append(quote); pos++; continue; }
                        closed = true; break;
                    }
                    if (!closed) throw new FormatException("公式引号未闭合。");
                    _tokens.Add(new(quote == '"' ? "text" : "quoted", text.ToString())); continue;
                }
                if (ch == '#')
                {
                    string[] errors = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!", "#CYCLE!", "#SCRIPT!", "#CALC!", "#SPILL!"];
                    var error = errors.FirstOrDefault(e => input.AsSpan(pos).StartsWith(e, StringComparison.OrdinalIgnoreCase));
                    if (error is null) { _tokens.Add(new("#", "#")); pos++; continue; }
                    _tokens.Add(new("error", error)); pos += error.Length; continue;
                }
                if (char.IsAsciiDigit(ch) || ch == '.' && pos + 1 < input.Length && char.IsAsciiDigit(input[pos + 1]))
                {
                    var start = pos++;
                    while (pos < input.Length && (char.IsAsciiDigit(input[pos]) || input[pos] == '.')) pos++;
                    if (pos < input.Length && input[pos] is 'e' or 'E')
                    {
                        pos++; if (pos < input.Length && input[pos] is '+' or '-') pos++;
                        while (pos < input.Length && char.IsAsciiDigit(input[pos])) pos++;
                    }
                    _tokens.Add(new("number", input[start..pos])); continue;
                }
                if (char.IsLetter(ch) || ch is '_' or '$' or '\\')
                {
                    var start = pos++;
                    while (pos < input.Length && (char.IsLetterOrDigit(input[pos]) || input[pos] is '_' or '$' or '.' or '\\')) pos++;
                    _tokens.Add(new("word", input[start..pos])); continue;
                }
                var op = ch.ToString(); pos++;
                if (pos < input.Length && (op is "<" or ">") && (input[pos] == '=' || op == "<" && input[pos] == '>')) op += input[pos++];
                if (op is not ("+" or "-" or "*" or "/" or "^" or "&" or "=" or "<" or ">" or "<=" or ">=" or "<>" or "%" or "(" or ")" or "," or ":" or "!"))
                    throw new FormatException($"不支持的公式符号：{op}。");
                _tokens.Add(new(op, op));
            }
            _tokens.Add(new("end", ""));
        }

        public FormulaNode Parse()
        {
            var result = Expression(0, 0);
            if (Current.Kind != "end") throw new FormatException("公式包含多余内容或不支持的引用。");
            // Also bound AST depth for long left-associative expressions.
            var pending = new Stack<(FormulaNode Node, int Depth)>(); pending.Push((result, 0));
            while (pending.TryPop(out var item))
            {
                if (item.Depth > 128) throw new FormatException("公式嵌套超过 128 层。");
                switch (item.Node)
                {
                    case BinaryNode b: pending.Push((b.Left, item.Depth + 1)); pending.Push((b.Right, item.Depth + 1)); break;
                    case UnaryNode u: pending.Push((u.Operand, item.Depth + 1)); break;
                    case FunctionNode f: foreach (var a in f.Arguments) pending.Push((a, item.Depth + 1)); break;
                }
            }
            return result;
        }

        private FormulaNode Expression(int precedence, int depth)
        {
            if (depth > 128 || ++_nodes > 2048) throw new FormatException("公式过长或嵌套过深。");
            var first = Take();
            FormulaNode left;
            if (first.Kind is "+" or "-") left = new UnaryNode(first.Kind, Expression(70, depth + 1));
            else if (first.Kind == "(") { left = Expression(0, depth + 1); Expect(")"); }
            else if (first.Kind == "text") left = new ValueNode(CellValue.FromText(first.Text));
            else if (first.Kind == "table") left = StructuredReferences.Parse(first.Text);
            else if (first.Kind == "error") { CellInput.TryFromErrorText(first.Text, out var error); left = new ValueNode(error); }
            else if (first.Kind == "number" && Current.Kind != "!")
                left = new ValueNode(CellValue.FromNumber(double.Parse(first.Text, NumberStyles.Float, CultureInfo.InvariantCulture)));
            else if (first.Kind is "word" or "quoted" or "number") left = Identifier(first, depth);
            else throw new FormatException("公式缺少表达式。");
            while (true)
            {
                if (Current.Kind == "%" && 60 >= precedence) { Take(); left = new UnaryNode("%", left); continue; }
                if (Current.Kind == "#" && 80 >= precedence)
                {
                    if (left is not (ReferenceNode { IsRange: false, WholeColumns: false } or NameNode or ValueNode { Value.Kind: CellValueKind.Error }))
                        throw new FormatException("溢出引用必须指向单个公式锚点。");
                    Take(); left = new UnaryNode("#", left); continue;
                }
                var priority = Current.Kind switch { "=" or "<>" or "<" or ">" or "<=" or ">=" => 10, "&" => 20, "+" or "-" => 30, "*" or "/" => 40, "^" => 50, _ => -1 };
                if (priority < precedence) break;
                var op = Take().Kind;
                left = new BinaryNode(op, left, Expression(priority + 1, depth + 1));
            }
            return left;
        }

        private FormulaNode Identifier(Token token, int depth)
        {
            string? sheet = null;
            if (Current.Kind == "!") { Take(); sheet = token.Text; token = Take(); }
            else if (token.Kind == "quoted") throw new FormatException("带引号的表名必须后接 !。");
            // Excel retains the sheet qualifier when a structural deletion removes the referenced address.
            if (sheet is not null && token.Kind == "error" && token.Text.Equals("#REF!", StringComparison.OrdinalIgnoreCase))
                return new ValueNode(CellValue.FromError(CellErrorCode.Reference));
            if (token.Kind != "word") throw new FormatException("工作表引用缺少地址。");
            if (Current.Kind == "table") return StructuredReferences.Parse(Take().Text, token.Text, sheet);
            if (Current.Kind == "(" && sheet is null)
            {
                Take(); var args = new List<FormulaNode>(); var functionName = ExcelFormulaNames.CanonicalName(token.Text);
                if (Current.Kind != ")")
                {
                    while (true)
                    {
                        if (Current.Kind is "," or ")" && (functionName is "OFFSET" or "INDIRECT" || DynamicArrayFunctions.IsFunction(functionName) || functionName == "XLOOKUP" && (args.Count == 0 || args.Count >= 3) || functionName == "XMATCH" && args.Count >= 2 || functionName == "SUBTOTAL" && args.Count == 0 ||
                            functionName is "LARGE" or "SMALL" && args.Count==1 || functionName=="RANK.EQ" && args.Count is 0 or 2))
                        {
                            if (++_nodes > 2048) throw new FormatException("公式过长。");
                            args.Add(new MissingNode());
                        }
                        else args.Add(Expression(0, depth + 1));
                        if (Current.Kind != ",") break;
                        Take();
                    }
                }
                Expect(")");
                if (functionName == "ANCHORARRAY" && args.Count == 1 && args[0] is ReferenceNode { IsRange: false, WholeColumns: false })
                    return new UnaryNode("#", args[0]);
                return new FunctionNode(functionName, args);
            }
            if (sheet is null && token.Text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return new ValueNode(CellValue.FromBoolean(true));
            if (sheet is null && token.Text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return new ValueNode(CellValue.FromBoolean(false));
            var first = Endpoint(token.Text);
            if (first is null) return new NameNode(token.Text, sheet);
            if (Current.Kind != ":")
            {
                // Bare identifiers are names, including A; legacy A whole-column spelling is intentionally not guessed.
                if (first.Value.Row < 0) return new NameNode(token.Text, sheet);
                var p = first.Value;
                return new ReferenceNode(sheet, p.Row, p.Col, p.Row, p.Col, p.Ca, p.Ra);
            }
            Take();
            var end = Take();
            if (Current.Kind == "!") { Take(); if (!string.Equals(end.Text, sheet, StringComparison.OrdinalIgnoreCase)) throw new FormatException("不支持跨工作表的三维区域。"); end = Take(); }
            var second = Endpoint(end.Text) ?? throw new FormatException("区域结束地址无效。");
            var a = first.Value; var b = second;
            if ((a.Row < 0) != (b.Row < 0) || b.Row < a.Row || b.Col < a.Col) throw new FormatException("区域端点顺序或维度无效。");
            return new ReferenceNode(sheet, Math.Max(0, a.Row), a.Col, a.Row < 0 ? WorkbookLimits.MaxRows - 1 : b.Row,
                b.Col, a.Ca, a.Ra, b.Ca, b.Ra, true, a.Row < 0);
        }

        private static (int Row, int Col, bool Ca, bool Ra)? Endpoint(string text)
        {
            var match = CellPattern.Match(text);
            var whole = !match.Success;
            if (whole) match = ColumnPattern.Match(text);
            if (!match.Success) return null;
            var column = 0;
            foreach (var c in match.Groups[2].Value.ToUpperInvariant()) column = column * 26 + c - 'A' + 1;
            var row = whole ? -1 : int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) - 1;
            if (whole && column > WorkbookLimits.MaxColumns) return null;
            if (column > WorkbookLimits.MaxColumns || row >= WorkbookLimits.MaxRows) throw new FormatException("公式引用超过支持的坐标范围。");
            return (row, column - 1, match.Groups[1].Length != 0, !whole && match.Groups[3].Length != 0);
        }
        private void Expect(string kind) { if (Current.Kind != kind) throw new FormatException($"公式缺少 {kind}。"); Take(); }
    }
}
