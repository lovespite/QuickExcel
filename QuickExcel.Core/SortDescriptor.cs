using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public enum SortBasis { Value, FillColor, FontColor }

/// <summary>Immutable, ordered sort keys. Text comparison is ordinal, never a claim of Excel locale collation.</summary>
public sealed record SortKey
{
    public int Column { get; }
    public bool Descending { get; }
    public bool CaseSensitive { get; }
    public SortBasis Basis { get; }
    public string? PreferredColor { get; }
    public ImmutableArray<string> CustomOrder { get; }

    public SortKey(int column, bool descending = false, bool caseSensitive = false, SortBasis basis = SortBasis.Value,
        string? preferredColor = null, IEnumerable<string>? customOrder = null)
    {
        Column = column; Descending = descending; CaseSensitive = caseSensitive; Basis = basis;
        PreferredColor = preferredColor; CustomOrder = customOrder?.Take(257).ToImmutableArray() ?? [];
    }

    internal void Validate()
    {
        if (!Enum.IsDefined(Basis)) throw new ArgumentException("未知排序依据。");
        if (Basis != SortBasis.Value && (PreferredColor is null || PreferredColor.Length != 8 || PreferredColor.Any(c => !Uri.IsHexDigit(c))))
            throw new ArgumentException("颜色排序需要 8 位 ARGB 颜色。");
        if (CustomOrder.Length > 256 || CustomOrder.Any(value => value is null || value.Length > 1024) ||
            CustomOrder.Distinct(CaseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase).Count() != CustomOrder.Length ||
            Basis != SortBasis.Value && !CustomOrder.IsEmpty)
            throw new ArgumentException("自定义顺序需要最多 256 个不重复文本，且仅适用于值排序。");
    }

    internal int Compare((CellValue Value, StyleDefinition Style) a, (CellValue Value, StyleDefinition Style) b)
    {
        if (Basis != SortBasis.Value)
        {
            bool Matches(StyleDefinition style) => string.Equals(Basis == SortBasis.FillColor ? style.FillColor : style.FontColor,
                PreferredColor, StringComparison.OrdinalIgnoreCase);
            var comparison = Matches(b.Style).CompareTo(Matches(a.Style));
            return Descending ? -comparison : comparison;
        }
        if (!CustomOrder.IsEmpty)
        {
            int Rank(CellValue value) => value.Kind == CellValueKind.Text
                ? CustomOrder.IndexOf(value.TextValue, 0, CustomOrder.Length, CaseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase)
                : -1;
            var ar = Rank(a.Value); var br = Rank(b.Value);
            // Listed values precede unlisted values in both directions; unlisted values use the normal type order.
            if (ar >= 0 || br >= 0)
            {
                if (ar < 0) return 1;
                if (br < 0) return -1;
                return Descending ? br.CompareTo(ar) : ar.CompareTo(br);
            }
        }
        return SortRangeCommand.Compare(a.Value, b.Value, Descending, CaseSensitive);
    }
}

public sealed class SortDescriptor
{
    public SheetId SheetId { get; }
    public CellPosition First { get; }
    public CellPosition Last { get; }
    public bool HasHeader { get; }
    public ImmutableArray<SortKey> Keys { get; }
    /// <summary>Expand the rectangle across adjacent data columns so the whole record moves with the key column.</summary>
    public bool ExpandColumns { get; }

    public SortDescriptor(SheetId sheetId, CellPosition first, CellPosition last, IEnumerable<SortKey> keys, bool hasHeader = false, bool expandColumns = false)
    {
        SheetId = sheetId; First = first; Last = last; HasHeader = hasHeader; ExpandColumns = expandColumns;
        Keys = keys.Take(33).ToImmutableArray();
    }
}
