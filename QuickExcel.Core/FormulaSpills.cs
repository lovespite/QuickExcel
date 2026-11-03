using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>The only input in a spill is its anchor formula. Other positions are derived results.</summary>
public sealed record FormulaSpill(CellAddress Anchor, FormulaArray Array)
{
    public CellRectangle Range => new(Anchor.Row, Anchor.Column, checked(Anchor.Row + Array.Rows - 1), checked(Anchor.Column + Array.Columns - 1));
}

/// <summary>Per-round allocation; neither a failed reservation nor cancellation mutates document input.</summary>
public sealed class FormulaSpillAllocator(WorkbookSnapshot snapshot)
{
    private readonly Dictionary<CellAddress, CellAddress> _owners = [];
    private readonly Dictionary<CellAddress, FormulaSpill> _spills = [];
    private readonly int _inputCells = snapshot.Sheets.Sum(s => s.Cells.Values.Count(c => c.SpillOwner is null));
    private int _newCells;
    public IReadOnlyDictionary<CellAddress, FormulaSpill> Spills => _spills;
    public IReadOnlyDictionary<CellAddress, CellAddress> Owners => _owners;

    public bool TryReserve(FormulaSpill spill, out string? reason, Action? check = null)
    {
        check ??= () => { };
        var anchor = spill.Anchor; var sheet = snapshot.GetSheet(anchor.SheetId);
        reason = null;
        if (_spills.ContainsKey(anchor)) throw new InvalidOperationException("同一公式不能重复分配溢出区域。");
        if (sheet.GetCell(anchor.Row, anchor.Column)?.Content.Kind != CellContentKind.Formula)
            throw new InvalidOperationException("溢出所有者必须是公式单元格。");
        var range = spill.Range;
        if (range.Bottom >= WorkbookLimits.MaxRows || range.Right >= WorkbookLimits.MaxColumns)
        { reason = "溢出区域超出工作表边界。"; return false; }
        bool Intersects(CellRectangle other) => range.Top <= other.Bottom && range.Bottom >= other.Top && range.Left <= other.Right && range.Right >= other.Left;
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(Intersects))
        { reason = "溢出区域与合并单元格重叠。"; return false; }
        if (sheet.Data.Tables.Any(table => Intersects(table.Range)))
        { reason = "溢出区域与 Excel 表格重叠。"; return false; }
        if (sheet.Data.Pivots.Any(pivot => Intersects(pivot.Range)))
        { reason = "溢出区域与透视结果重叠。"; return false; }
        if ((long)_owners.Count + spill.Array.Values.Length > WorkbookLimits.MaxCells)
        { reason = "本轮溢出结果超过工作簿单元格容量。"; return false; }
        var additions = 0;
        for (var row = range.Top; row <= range.Bottom; row++)
        for (var column = range.Left; column <= range.Right; column++)
        {
            check(); var address = new CellAddress(anchor.SheetId, row, column);
            if (address == anchor) continue;
            var cell = sheet.GetCell(row, column);
            if (cell is null || cell.SpillOwner is not null) additions++;
            if (_owners.ContainsKey(address) || cell is { SpillOwner: null } &&
                (cell.Content.Kind == CellContentKind.Formula || cell.Content.LiteralValue.Kind != CellValueKind.Blank))
            { reason = "溢出区域被已有内容或其他数组占用。"; return false; }
        }
        if ((long)_inputCells + _newCells + additions > WorkbookLimits.MaxCells)
        { reason = "溢出结果与已有输入合计超过工作簿单元格容量。"; return false; }
        // All checks precede admission; cancelled allocation is never partially observable.
        check();
        for (var row = range.Top; row <= range.Bottom; row++)
        for (var column = range.Left; column <= range.Right; column++)
            _owners.Add(new(anchor.SheetId, row, column), anchor);
        _spills.Add(anchor, spill);
        _newCells += additions;
        return true;
    }

    public ImmutableDictionary<CellAddress, FormulaSpill> Snapshot() => _spills.ToImmutableDictionary();
}
