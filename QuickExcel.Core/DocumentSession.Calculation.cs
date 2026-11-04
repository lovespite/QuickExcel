using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed partial class DocumentSession
{
    // Shared validation/application for published rounds and private action drafts.
    // This method has no session, history, notification or script-event side effects.
    internal static WorkbookSnapshot ApplyCalculation(WorkbookSnapshot snapshot, CalculationResult result, out ImmutableArray<CellChange> cellChanges)
    {
        var sheets = snapshot.Sheets.ToArray();
        var builders = new Dictionary<SheetId, ImmutableDictionary<CellPosition, CellSnapshot>.Builder>();
        var changes = ImmutableArray.CreateBuilder<CellChange>();
        var allocation = new FormulaSpillAllocator(snapshot);
        foreach (var (address, calculation) in result.Cells.OrderBy(p => p.Key.SheetId.Value).ThenBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
            if (calculation.Array is { } array)
            {
                if (calculation.State is not (CalculationState.Fresh or CalculationState.Error) ||
                    !allocation.TryReserve(new(address, array), out _))
                    throw new InvalidOperationException("数组结果未通过溢出区域校验，整轮结果未提交。");
            }
        foreach (var sheet in sheets)
        foreach (var (position, cell) in sheet.Cells)
            if (cell.SpillOwner is { } owner && result.Cells.TryGetValue(owner, out var ownerResult))
            {
                if (!builders.TryGetValue(sheet.Id, out var builder)) builders[sheet.Id] = builder = sheet.Cells.ToBuilder();
                if (ownerResult.State == CalculationState.Blocked)
                {
                    var retained = new CellSnapshot(cell.Content, new(cell.Calculation!.Value, CalculationState.Blocked,
                        result.Stamp.InputRevision, result.Stamp.CalculationGeneration, result.Stamp.ScriptGeneration, ownerResult.Diagnostics), owner);
                    builder[position] = retained; changes.Add(new(new(sheet.Id,position.Row,position.Column),cell,retained)); continue;
                }
                var after = cell.Content.StyleId == sheet.GetInheritedStyleId(position.Row, position.Column) ? null : new CellSnapshot(cell.Content);
                if (after is null) builder.Remove(position); else builder[position] = after;
                changes.Add(new(new(sheet.Id, position.Row, position.Column), cell, after));
            }
        foreach (var (address, calculation) in result.Cells)
        {
            WorkbookLimits.ValidatePosition(address.Row, address.Column);
            var sheet = snapshot.GetSheet(address.SheetId);
            var before = sheet.GetCell(address.Row, address.Column);
            if (before?.Content.Kind != CellContentKind.Formula)
                throw new InvalidOperationException("计算结果只能提交到本轮快照的公式单元格。");
            if (calculation.InputRevision != result.Stamp.InputRevision || calculation.CalculationGeneration != result.Stamp.CalculationGeneration || calculation.ScriptGeneration != result.Stamp.ScriptGeneration)
                throw new InvalidOperationException("计算单元格的修订或代次不匹配。");
            if (calculation.Value is { Kind: CellValueKind.Text } text && text.TextValue.Length > WorkbookLimits.MaxCellTextLength)
                throw new InvalidOperationException("计算结果超过单元格文本容量。");
            if (snapshot.FileState?.FormulaReadOnly == true && (calculation.State != CalculationState.Blocked || calculation.Value != before.Calculation?.Value))
                throw new InvalidOperationException("只读公式工作簿只能保留原缓存并标记受阻，不能发布新结果。");
            if (!builders.TryGetValue(sheet.Id, out var builder)) builders[sheet.Id] = builder = sheet.Cells.ToBuilder();
            var committed = calculation.State == CalculationState.Blocked && before.Calculation?.Array is { } oldArray && calculation.Value == oldArray[0,0]
                ? new CellCalculation(calculation.Value, calculation.State, calculation.InputRevision, calculation.CalculationGeneration,
                    calculation.ScriptGeneration, calculation.Diagnostics, oldArray, calculation.BlockedSpill) : calculation;
            var content = calculation.Array is not null && !before.Content.DynamicArray
                ? CellContent.FromFormula(before.Content.FormulaText, before.Content.StyleId, dynamicArray: true) : before.Content;
            var after = new CellSnapshot(content, committed);
            builder[new(address.Row, address.Column)] = after;
            changes.Add(new(address, before, after));
        }
        foreach (var spill in allocation.Spills.Values)
        {
            var sheet = snapshot.GetSheet(spill.Anchor.SheetId); var builder = builders[sheet.Id];
            for (var row = 0; row < spill.Array.Rows; row++)
            for (var column = 0; column < spill.Array.Columns; column++)
            {
                if (row == 0 && column == 0) continue;
                var position = new CellPosition(spill.Anchor.Row + row, spill.Anchor.Column + column);
                var before = sheet.GetCell(position.Row, position.Column); var value = spill.Array[row, column];
                var calculation = new CellCalculation(value, value.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh,
                    result.Stamp.InputRevision, result.Stamp.CalculationGeneration, result.Stamp.ScriptGeneration);
                var after = new CellSnapshot(CellContent.FromLiteral(CellValue.Blank, sheet.GetStyleId(position.Row, position.Column)), calculation, spill.Anchor);
                builder[position] = after;
                changes.Add(new(new(sheet.Id, position.Row, position.Column), before, after));
            }
        }
        for (var index = 0; index < sheets.Length; index++)
            if (builders.TryGetValue(sheets[index].Id, out var builder)) sheets[index] = sheets[index].WithCells(builder.ToImmutable());
        cellChanges = changes.ToImmutable();
        return PivotSafety.MarkSourceChanges(snapshot, snapshot.WithSheets(sheets));
    }
}
