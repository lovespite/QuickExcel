using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal enum SearchScope { Sheet, Workbook, Selection }
internal enum SearchTarget { Display, Formula }
internal sealed record SearchOptions(string Query, SearchScope Scope = SearchScope.Sheet, SearchTarget Target = SearchTarget.Display,
    bool MatchCase = false, bool WholeCell = false, bool IncludeHidden = false);
internal sealed record SearchHit(CellAddress Address, int SheetIndex, string SheetName, string Text, CellContent Content)
{
    public string Location => $"{SheetName}!{Column(Address.Column)}{Address.Row + 1}";
    private static string Column(int value) { var text = ""; do { text = (char)('A' + value % 26) + text; value = value / 26 - 1; } while (value >= 0); return text; }
}

internal static class WorkbookSearch
{
    public const int Limit = 20_000;
    public static IReadOnlyList<SearchHit> Find(WorkbookSnapshot snapshot, int activeSheet, SearchOptions options,
        IReadOnlySet<CellAddress>? selection = null, Func<int, int, bool>? filtered = null)
    {
        if (string.IsNullOrEmpty(options.Query)) throw new ArgumentException("请输入查找内容。");
        if (options.Query.Length > WorkbookLimits.MaxCellTextLength) throw new ArgumentException("查找内容过长。");
        if (options.Scope == SearchScope.Selection && (selection is null || selection.Count == 0)) throw new ArgumentException("请先选择查找区域。");
        var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var result = new List<SearchHit>();
        for (var index = 0; index < snapshot.Sheets.Length; index++)
        {
            if (options.Scope != SearchScope.Workbook && index != activeSheet) continue;
            var sheet = snapshot.Sheets[index];
            if (sheet.Visibility != SheetVisibility.Visible) continue;
            foreach (var (position, cell) in sheet.Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
            {
                var address = new CellAddress(sheet.Id, position.Row, position.Column);
                if (options.Scope == SearchScope.Selection && !selection!.Contains(address)) continue;
                if (!options.IncludeHidden && (sheet.Rows.GetValueOrDefault(position.Row)?.Hidden == true ||
                    sheet.Columns.GetValueOrDefault(position.Column)?.Hidden == true || filtered?.Invoke(index, position.Row) == true)) continue;
                if (options.Target == SearchTarget.Formula && cell.Content.Kind != CellContentKind.Formula) continue;
                var text = options.Target == SearchTarget.Formula ? cell.Content.FormulaText :
                    CellDisplayFormatter.Format(cell.Value, snapshot.Styles[sheet.GetStyleId(position.Row, position.Column)], snapshot.DateSystem);
                if (!(options.WholeCell ? text.Equals(options.Query, comparison) : text.Contains(options.Query, comparison))) continue;
                if (result.Count == Limit) throw new InvalidOperationException("匹配超过 20,000 格，请缩小查找范围。");
                result.Add(new(address, index, sheet.Name, text, cell.Content));
            }
        }
        return result;
    }

    public static SetCellsCommand Replace(IReadOnlyList<SearchHit> hits, SearchOptions options, string replacement)
    {
        if (string.IsNullOrEmpty(options.Query)) throw new ArgumentException("请输入查找内容。");
        if (hits.Count > Limit || replacement.Length > WorkbookLimits.MaxCellTextLength) throw new ArgumentException("替换超出容量。");
        var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var edits = new List<CellEdit>();
        foreach (var hit in hits)
        {
            if (options.Target == SearchTarget.Display && hit.Content.Kind == CellContentKind.Formula)
                throw new InvalidOperationException("匹配包含公式结果。请改为搜索公式源码，或缩小范围；本次没有替换。");
            // Bound expansion before String.Replace allocates the resulting string.
            if (!options.WholeCell && replacement.Length > options.Query.Length)
            {
                long length = hit.Text.Length;
                var start = 0;
                while ((start = hit.Text.IndexOf(options.Query, start, comparison)) >= 0)
                {
                    length += replacement.Length - options.Query.Length;
                    if (length > WorkbookLimits.MaxCellTextLength) throw new ArgumentException("替换后的单元格文本超出容量；本次没有替换。");
                    start += options.Query.Length;
                }
            }
            var text = options.WholeCell ? replacement : hit.Text.Replace(options.Query, replacement, comparison);
            WorkbookLimits.ValidateText(text);
            CellContent content;
            if (options.Target == SearchTarget.Formula)
            {
                if (!text.StartsWith('=') || text.Length < 2) throw new ArgumentException("替换后必须仍为有效公式。");
                QuickExcel.Formulas.FormulaSyntax.Parse(text);
                content = CellContent.FromFormula(text, hit.Content.StyleId);
            }
            else if (hit.Content.LiteralValue.Kind == CellValueKind.Text)
                content = CellContent.FromLiteral(CellValue.FromText(text), hit.Content.StyleId);
            else
            {
                content = CellInput.Parse(text, hit.Content.StyleId);
                if (content.Kind == CellContentKind.Formula)
                    throw new ArgumentException("值替换不能创建公式；请在公式栏显式输入。");
            }
            edits.Add(new(hit.Address, content));
        }
        return new(edits);
    }
}
