using System.Collections;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public readonly record struct GridRange(int Top, int Left, int Bottom, int Right)
{
    public int Rows => Bottom - Top + 1;
    public int Columns => Right - Left + 1;
    public long Count => (long)Rows * Columns;
    public static GridRange Between(CellPosition a, CellPosition b) =>
        new(Math.Min(a.Row, b.Row), Math.Min(a.Column, b.Column), Math.Max(a.Row, b.Row), Math.Max(a.Column, b.Column));

    public static GridRange Rectangle(IEnumerable<CellPosition> positions, Func<int, bool>? excludedRow = null)
    {
        var cells = positions.Take(100_001).ToHashSet();
        if (cells.Count == 0) throw new ArgumentException(UiText.Get("请先选择单元格。"));
        if (cells.Count > 100_000) throw new ArgumentException(UiText.Get("复制范围上限为 100,000 格。"));
        var range = new GridRange(cells.Min(c => c.Row), cells.Min(c => c.Column), cells.Max(c => c.Row), cells.Max(c => c.Column));
        var expected = excludedRow is null ? range.Count : (long)Enumerable.Range(range.Top, range.Rows).Count(r => !excludedRow(r)) * range.Columns;
        var actual = excludedRow is null ? cells.Count : cells.Count(c => !excludedRow(c.Row));
        if (expected != actual) throw new ArgumentException(UiText.Get("请使用连续矩形选区；不连续选区不能复制或粘贴。"));
        return range;
    }
}

/// <summary>Bounded TSV exchange. Normal paste interprets numbers/booleans but keeps formula text literal.</summary>
public static class GridClipboard
{
    public const int MaxCharacters = 4 * 1024 * 1024;
    public const int MaxEditCells = 20_000;

    public static string Write(GridRange range, Func<int, int, string> value, Func<int, bool>? excludedRow = null)
    {
        if (range.Count > 100_000) throw new ArgumentException(UiText.Get("复制范围上限为 100,000 格。"));
        var result = new StringBuilder();
        var wroteRow = false;
        for (var row = range.Top; row <= range.Bottom; row++)
        {
            if (excludedRow?.Invoke(row) == true) continue;
            if (wroteRow) result.Append("\r\n");
            wroteRow = true;
            for (var column = range.Left; column <= range.Right; column++)
            {
                if (column > range.Left) result.Append('\t');
                var text = value(row, column);
                if (text.IndexOfAny(['\t', '\r', '\n', '"']) >= 0)
                    result.Append('"').Append(text.Replace("\"", "\"\"")).Append('"');
                else result.Append(text);
                if (result.Length > MaxCharacters) throw new ArgumentException(UiText.Get("复制文本不能超过 4 Mi 个字符。"));
            }
        }
        return result.ToString();
    }

    public static string[][] Read(string text)
    {
        if (text.Length > MaxCharacters) throw new ArgumentException(UiText.Get("粘贴文本不能超过 4 Mi 个字符。"));
        var rows = QuickExcel.Core.DelimitedText.ReadRows(text, '\t', MaxEditCells).ToArray();
        if (rows.Any(row => row.Length != rows[0].Length)) throw new ArgumentException(UiText.Get("粘贴数据必须是矩形，每行列数应相同。"));
        return rows;
    }

    public static (int Row, int Column, string Value)[] Prepare(string text, GridRange target, bool allowFormulas)
    {
        var data = Read(text);
        var fill = data.Length == 1 && data[0].Length == 1;
        var rows = fill ? target.Rows : data.Length;
        var columns = fill ? target.Columns : data[0].Length;
        if (target.Count != 1 && !fill && (target.Rows != rows || target.Columns != columns))
            throw new ArgumentException(UiText.Get("选区尺寸须与粘贴数据相同，或只选择左上角单元格。"));
        if ((long)rows * columns > MaxEditCells || target.Top < 0 || target.Left < 0 ||
            target.Top + rows > WorkbookLimits.MaxRows || target.Left + columns > WorkbookLimits.MaxColumns)
            throw new ArgumentException(UiText.Format($"粘贴超出 20,000 格或界面 {WorkbookLimits.MaxRows:N0} 行 × {WorkbookLimits.MaxColumns:N0} 列范围。"));
        var edits = new List<(int, int, string)>();
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < columns; c++)
            {
                var input = data[fill ? 0 : r][fill ? 0 : c];
                // Apostrophes copied as data must not disappear. Formula interpretation is opt-in.
                if (input.StartsWith('\'') || (!allowFormulas && input.StartsWith('='))) input = "'" + input;
                CellInput.Parse(input); // Validate the whole patch before the command path sees it.
                edits.Add((target.Top + r, target.Left + c, input));
            }
        return edits.ToArray();
    }
}

public sealed class CellEditState
{
    public sealed record FormulaDraft(int Sheet, int Row, int Column);
    public FormulaDraft? Formula { get; private set; }
    public bool GridEditing { get; private set; }
    public bool IsComposing { get; private set; }
    public void BeginFormula(int sheet, int row, int column) => Formula ??= new(sheet, row, column);
    public void EndFormula() => Formula = null;
    public void BeginGrid() => GridEditing = true;
    public void EndGrid() => GridEditing = false;
    public void StartComposition() => IsComposing = true;
    public void EndComposition() => IsComposing = false;
    public void Reset() { Formula = null; GridEditing = IsComposing = false; }
}

/// <summary>Stable row identities without allocating one object per logical row at bind time.</summary>
public sealed class VirtualRowList<T>(int count, Func<int, T> factory, Func<T, int> index) : IList where T : class
{
    private readonly Dictionary<int, T> _created = [];
    public int Count { get; } = count;
    public int CreatedCount => _created.Count;
    public T this[int i] => i < 0 || i >= Count ? throw new ArgumentOutOfRangeException(nameof(i)) :
        _created.TryGetValue(i, out var value) ? value : _created[i] = factory(i);
    object? IList.this[int i] { get => this[i]; set => throw new NotSupportedException(); }
    public int IndexOf(object? value) => value is T row && _created.TryGetValue(index(row), out var stored) && ReferenceEquals(row, stored) ? index(row) : -1;
    public bool Contains(object? value) => IndexOf(value) >= 0;
    public bool IsReadOnly => true;
    public bool IsFixedSize => true;
    public bool IsSynchronized => false;
    public object SyncRoot => this;
    public IEnumerator GetEnumerator() { for (var i = 0; i < Count; i++) yield return this[i]; }
    public void CopyTo(Array array, int offset) { for (var i = 0; i < Count; i++) array.SetValue(this[i], i + offset); }
    public int Add(object? value) => throw new NotSupportedException();
    public void Clear() => throw new NotSupportedException();
    public void Insert(int index, object? value) => throw new NotSupportedException();
    public void Remove(object? value) => throw new NotSupportedException();
    public void RemoveAt(int index) => throw new NotSupportedException();
}
