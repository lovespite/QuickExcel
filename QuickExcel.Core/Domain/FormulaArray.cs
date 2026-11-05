using QuickExcel.Core;
using System.Collections.Immutable;

namespace QuickExcel.Domain;

public readonly record struct FormulaArrayDimensions(int Rows, int Columns);

/// <summary>A bounded, immutable rectangular formula value. It is a result, never cell input.</summary>
public sealed class FormulaArray
{
    public const int MaxElements = 100_000;
    public const int MaxTextCharacters = 4 * 1024 * 1024;

    public FormulaArray(int rows, int columns, IEnumerable<CellValue> values)
    {
        ValidateShape(rows, columns);
        ArgumentNullException.ThrowIfNull(values);
        Rows = rows; Columns = columns;
        var count = checked(rows * columns);
        var builder = ImmutableArray.CreateBuilder<CellValue>(count);
        var characters = 0;
        foreach (var value in values)
        {
            if (builder.Count == count) throw new ArgumentException(MessageText.Get("数组元素数量与形状不一致。"), nameof(values));
            if (value.Kind == CellValueKind.Text)
            {
                WorkbookLimits.ValidateText(value.TextValue);
                characters = checked(characters + value.TextValue.Length);
                if (characters > MaxTextCharacters) throw new ArgumentException(MessageText.Get("单个公式数组文本超过 4 Mi 字符。"), nameof(values));
            }
            builder.Add(value);
        }
        if (builder.Count != count) throw new ArgumentException(MessageText.Get("数组元素数量与形状不一致。"), nameof(values));
        Values = builder.MoveToImmutable();
    }

    public int Rows { get; }
    public int Columns { get; }
    public ImmutableArray<CellValue> Values { get; }
    public CellValue this[int row, int column] => row >= 0 && row < Rows && column >= 0 && column < Columns
        ? Values[row * Columns + column] : throw new ArgumentOutOfRangeException(nameof(row));

    public static void ValidateShape(int rows, int columns)
    {
        if (rows < 1 || rows > WorkbookLimits.MaxRows || columns < 1 || columns > WorkbookLimits.MaxColumns || (long)rows * columns > MaxElements)
            throw new ArgumentOutOfRangeException(nameof(rows), MessageText.Get("公式数组必须非空，且不能超过工作表边界或 100,000 个元素。"));
    }

    public bool ContentEquals(FormulaArray other) => Rows == other.Rows && Columns == other.Columns && Values.SequenceEqual(other.Values);
}
