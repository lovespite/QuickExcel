using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Bounded example synthesis. Only literal examples train the rule; only blank targets are filled.</summary>
public static class FlashFill
{
    private sealed record Atom(Func<int, string?> Read, int Cost);
    private sealed record Part(Atom? Atom, string Literal);
    private sealed record Pattern(Part[] Parts, int Cost)
    {
        public string? Read(int row)
        {
            var result = new System.Text.StringBuilder();
            foreach (var part in Parts)
            {
                var value = part.Atom is null ? part.Literal : part.Atom.Read(row);
                if (value is null) return null;
                result.Append(value);
            }
            return result.ToString();
        }
    }

    public static CleaningPreview Preview(WorkbookSnapshot book, SheetId sheetId, CellRectangle records, int targetColumn)
    {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        void CheckBudget() { if (timer.Elapsed > TimeSpan.FromSeconds(5)) throw new InvalidOperationException(MessageText.Get("快速填充推断超过 5 秒，请缩小区域或补充更清晰的示例。")); }
        var sheet = DataCleaning.RequireRange(book, sheetId, records);
        if (records.Columns is < 2 or > 17 || targetColumn < records.Left || targetColumn > records.Right || records.Count > DataCleaning.MaxCells)
            throw new ArgumentException(MessageText.Get("快速填充需要包含示例列及 1–16 列来源的数据区域（最多 20,000 格）。"));
        string Text(int row, int column)
        {
            var cell = sheet.GetCell(row, column);
            if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation is not { State: CalculationState.Fresh } || cell.Calculation.InputRevision != book.Revision))
                throw new InvalidOperationException(MessageText.Get("快速填充来源含待重算或错误公式，请先重算。"));
            var value = CellInput.DisplayText(cell?.Value ?? CellValue.Blank);
            if (value.Length > 256) throw new ArgumentException(MessageText.Get("快速填充每格最多处理 256 个字符，请缩小文本。"));
            return value;
        }
        var sources = Enumerable.Range(records.Left, records.Columns).Where(c => c != targetColumn).ToArray();
        if (sheet.Cells.Where(p => records.Contains(p.Key.Row, p.Key.Column)).Sum(p => (long)CellInput.DisplayText(p.Value.Value).Length) > 256_000)
            throw new ArgumentException(MessageText.Get("快速填充区域的文本总量超过 256,000 字符，请缩小区域。"));
        var rows = Enumerable.Range(records.Top, records.Rows).Where(r => sources.Any(c => Text(r, c).Length > 0)).ToArray();
        var examples = rows.Where(r => sheet.GetCell(r, targetColumn)?.Content is { Kind: CellContentKind.Literal, LiteralValue.Kind: not CellValueKind.Blank }).ToArray();
        var blanks = rows.Where(r => sheet.GetCell(r, targetColumn) is null || sheet.GetCell(r, targetColumn)?.Content is { Kind: CellContentKind.Literal, LiteralValue.Kind: CellValueKind.Blank }).ToArray();
        if (examples.Length == 0) throw new ArgumentException(MessageText.Get("请先在目标列输入至少一个示例，再按 Ctrl+E。选择数据行时请排除标题。"));
        if (blanks.Length == 0) throw new ArgumentException(MessageText.Get("目标列没有可填充的空白单元格。"));
        var first = examples[0]; var sample = Text(first, targetColumn); var atoms = new List<Atom>();
        foreach (var column in sources)
        {
            var cache = rows.ToDictionary(r => r, r => Text(r, column));
            string Value(int row) => cache[row];
            var readers = new List<(Func<int, string?> Read, int Cost)> { (Value, 1), (r => Value(r).Trim(), 2) };
            foreach (var regex in new[] { @"[\p{L}\p{M}\p{N}]+", @"\p{L}+", @"\d+", @"[^,;\t@/\\|\s]+" })
            {
                var tokens = cache.ToDictionary(p => p.Key, p => Regex.Matches(p.Value, regex).Select(m => m.Value).Take(32).ToArray());
                for (var i = 0; i < Math.Min(16, tokens[first].Length); i++)
                {
                    var index = i;
                    readers.Add((r => tokens[r].Length > index ? tokens[r][index] : null, 2));
                    readers.Add((r => tokens[r].Length > index ? tokens[r][^(index + 1)] : null, 2));
                }
            }
            for (var length = 1; length <= Math.Min(32, Value(first).Length); length++)
            {
                var size = length;
                readers.Add((r => Value(r).Length >= size && (Value(r).Length == size || !char.IsLowSurrogate(Value(r)[size])) ? Value(r)[..size] : null, 5));
                readers.Add((r => Value(r).Length >= size && !char.IsLowSurrogate(Value(r)[^size]) ? Value(r)[^size..] : null, 5));
            }
            foreach (var (read, cost) in readers)
            {
                atoms.Add(new(read, cost));
                atoms.Add(new(r => read(r)?.ToLowerInvariant(), cost + 1));
                atoms.Add(new(r => read(r)?.ToUpperInvariant(), cost + 1));
            }
        }
        // Equivalent atoms on every data row cannot introduce prediction ambiguity.
        string Signature(Atom atom)
        {
            CheckBudget();
            using var hash = System.Security.Cryptography.IncrementalHash.CreateHash(System.Security.Cryptography.HashAlgorithmName.SHA256);
            foreach (var row in rows)
            {
                var value = atom.Read(row);
                hash.AppendData(BitConverter.GetBytes(value?.Length ?? -1));
                if (value is not null) hash.AppendData(System.Text.Encoding.UTF8.GetBytes(value));
            }
            return Convert.ToHexString(hash.GetHashAndReset());
        }
        atoms = atoms.Where(a => !string.IsNullOrEmpty(a.Read(first)))
            .GroupBy(Signature)
            .Select(g => g.MinBy(a => a.Cost)!).ToList();
        var paths = new List<Pattern>[sample.Length + 1];
        for (var i = 0; i < paths.Length; i++) paths[i] = [];
        paths[0].Add(new([], 0));
        for (var offset = 0; offset < sample.Length; offset++)
        {
            CheckBudget();
            var candidates = paths[offset].OrderBy(p => p.Cost).Take(128).ToArray();
            foreach (var path in candidates)
            {
                if (path.Parts.Length >= 12) continue;
                foreach (var atom in atoms)
                {
                    var value = atom.Read(first)!;
                    if (sample.AsSpan(offset).StartsWith(value, StringComparison.Ordinal))
                        paths[offset + value.Length].Add(new([.. path.Parts, new(atom, "")], path.Cost + atom.Cost));
                }
                var last = path.Parts.LastOrDefault();
                var parts = last is { Atom: null } ? path.Parts[..^1].Append(new(null, last.Literal + sample[offset])).ToArray()
                    : [.. path.Parts, new Part(null, sample[offset].ToString())];
                paths[offset + 1].Add(new(parts, path.Cost + (char.IsLetterOrDigit(sample[offset]) ? 10 : 1)));
            }
            // Bound intermediate search even for repetitive examples.
            for (var i = offset + 1; i < paths.Length; i++)
                if (paths[i].Count > 512) paths[i] = paths[i].OrderBy(p => p.Cost).Take(256).ToList();
        }
        var matching = paths[^1].Where(p => p.Parts.Any(a => a.Atom is not null) && examples.All(r => p.Read(r) == Text(r, targetColumn))
            && blanks.All(r => p.Read(r) is not null)).OrderBy(p => p.Cost).ToArray();
        if (matching.Length == 0) throw new InvalidOperationException(MessageText.Get("无法从现有示例识别一致的填充规则。请补充示例，或选择不含标题的数据行。"));
        var best = matching[0];
        if (matching.Where(p => p.Cost == best.Cost).Any(p => blanks.Any(r => p.Read(r) != best.Read(r))))
            throw new InvalidOperationException(MessageText.Get("示例有多种可能的解释。请再输入一个能区分规则的示例，然后重试快速填充。"));
        var edits = blanks.Select(r => new CellEdit(new(sheetId, r, targetColumn),
            CellContent.FromLiteral(CellValue.FromText(best.Read(r)!), sheet.GetStyleId(r, targetColumn))));
        return new(book, new SetCellsCommand(edits).Prepare(book), [], [],
            [$"依据 {examples.Length} 个示例填充 {blanks.Length} 个空白格。请核对预览；保留所有示例和公式，结果为文本，可整体撤销。"]);
    }
}
