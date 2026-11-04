using System.Buffers.Binary;
using System.Collections.Immutable;
using OpenMcdf;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Clipboard;

/// <summary>
/// A narrowly scoped companion to SpreadsheetML: original range and raw numeric values only.
/// Never interprets BIFF formulas, strings, macros, embedded objects or external references.
/// Excel omits DATE1904 from clipboard BIFF; numeric/date pairs disambiguate its epoch instead.
/// </summary>
public sealed record Biff8ClipboardNumbers(int Top, int Left, int Rows, int Columns, ImmutableDictionary<CellPosition, double> Numbers)
{
    public static Biff8ClipboardNumbers Read(byte[] bytes)
    {
        if (bytes.Length < 512 || bytes.Length > SpreadsheetClipboardCodec.MaxBytes) throw new InvalidDataException("Biff8 剪贴板大小无效。");
        using var input = new MemoryStream(bytes, writable: false);
        using var cfb = RootStorage.Open(input, StorageModeFlags.LeaveOpen);
        using var workbook = cfb.OpenStream("Workbook");
        if (workbook.Length is < 20 or > SpreadsheetClipboardCodec.MaxBytes) throw new InvalidDataException("Biff8 Workbook 流超过限额。");
        var data = new byte[(int)workbook.Length]; workbook.ReadExactly(data);
        ushort U16(int offset) => BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset, 2));
        uint U32(int offset) => BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(offset, 4));
        double Number(int offset) => BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(data.AsSpan(offset, 8)));
        static double Rk(uint encoded)
        {
            var number = (encoded & 2) != 0 ? (int)encoded >> 2 : BitConverter.Int64BitsToDouble((long)(encoded & 0xFFFFFFFC) << 32);
            return (encoded & 1) != 0 ? number / 100 : number;
        }
        var count = 0; int? sheetOffset = null; var globals = true; var finished = false; var top = -1; var left = -1; var rows = 0; var columns = 0;
        var numbers = ImmutableDictionary.CreateBuilder<CellPosition, double>();
        void Add(int row, int col, double number)
        {
            if (!double.IsFinite(number) || top < 0 || row < top || col < left || row >= top + rows || col >= left + columns || !numbers.TryAdd(new(row - top, col - left), number))
                throw new InvalidDataException("Biff8 数值记录重复、越界或无效。");
        }
        for (var offset = 0; offset < data.Length;)
        {
            if (data.Length - offset < 4 || ++count > SpreadsheetClipboardCodec.MaxNodes) throw new InvalidDataException("Biff8 记录超过限额或被截断。");
            var id = U16(offset); var length = U16(offset + 2); var p = offset + 4; var end = checked(p + length);
            if (end > data.Length || length > 8224) throw new InvalidDataException("Biff8 记录长度无效。");
            if (offset == 0 || offset == sheetOffset)
            {
                if (id != 0x0809 || length != 16 || U16(p) != 0x0600 || U16(p + 2) != (globals ? 5 : 16)) throw new InvalidDataException("不是已验证的 BIFF8 工作簿/工作表流。");
            }
            if (id == 0x002F) throw new NotSupportedException("加密 Biff8 剪贴板不受支持。");
            if (globals && id == 0x0085)
            {
                if (sheetOffset is not null || length < 8) throw new NotSupportedException("Biff8 剪贴板需要单个工作表。");
                sheetOffset = checked((int)U32(p));
            }
            if (id == 0x000A)
            {
                if (!globals) { finished = true; break; }
                if (sheetOffset is null || sheetOffset < end || sheetOffset >= data.Length) throw new InvalidDataException("Biff8 工作表偏移无效。");
                globals = false; offset = sheetOffset.Value; continue;
            }
            if (!globals)
            {
                if (id == 0x0200)
                {
                    if (length != 14 || top >= 0) throw new InvalidDataException("Biff8 尺寸记录无效。");
                    top = checked((int)U32(p)); left = U16(p + 8); rows = checked((int)U32(p + 4) - top); columns = U16(p + 10) - left;
                    if (top < 0 || left < 0 || rows < 1 || columns < 1 || (long)rows * columns > WorkbookClipboard.MaxCells || top + (long)rows > 65536 || left + (long)columns > 256)
                        throw new NotSupportedException("Biff8 辅助数值范围不可表示或超过剪贴板限额。");
                }
                else if (id == 0x0203) { if (length != 14) throw new InvalidDataException("NUMBER 长度无效。"); Add(U16(p), U16(p + 2), Number(p + 6)); }
                else if (id == 0x027E) { if (length != 10) throw new InvalidDataException("RK 长度无效。"); Add(U16(p), U16(p + 2), Rk(U32(p + 6))); }
                else if (id == 0x00BD)
                {
                    if (length < 12 || (length - 6) % 6 != 0 || U16(end - 2) - U16(p + 2) + 1 != (length - 6) / 6) throw new InvalidDataException("MULRK 长度无效。");
                    for (var i = 0; i < (length - 6) / 6; i++) Add(U16(p), U16(p + 2) + i, Rk(U32(p + 6 + 6 * i)));
                }
                else if (id == 0x0006)
                {
                    if (length < 22) throw new InvalidDataException("FORMULA 缓存长度无效。");
                    if (U16(p + 12) != 0xFFFF) Add(U16(p), U16(p + 2), Number(p + 6));
                }
            }
            offset = end;
        }
        if (!finished || top < 0) throw new InvalidDataException("Biff8 缺少完整的工作表流。");
        return new(top, left, rows, columns, numbers.ToImmutable());
    }
}
