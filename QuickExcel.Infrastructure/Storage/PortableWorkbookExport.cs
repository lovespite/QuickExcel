using System.Collections.Immutable;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using static QuickExcel.Storage.PackageIO;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage : IPortableWorkbookStorage
{
    public PortableExportReport InspectPortableExport(WorkbookSnapshot snapshot)
    {
        if (snapshot.FileState?.FormulaReadOnly == true)
        {
            var formulas = snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula).ToArray();
            return new(0, formulas.Length, formulas.Count(c => c.Calculation?.Value is null), formulas.Count(c => c.Calculation?.Value is not null));
        }
        var index = PortableIndex(snapshot);
        var frozen = index.NonPortableCells.Select(a => snapshot.GetSheet(a.SheetId).GetCell(a.Row, a.Column)!).ToArray();
        return new(index.Nodes.Count - frozen.Length, frozen.Length,
            frozen.Count(c => c.Calculation?.Value is null),
            frozen.Count(c => c.Calculation?.Value is not null && (c.Calculation.InputRevision != snapshot.Revision
                || c.Calculation.State is not (CalculationState.Fresh or CalculationState.Error))));
    }

    public void ExportPortable(string path, WorkbookSnapshot snapshot, bool allowStaleCaches = false, bool discardUnsupported = false)
    {
        if (snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException("只读公式工作簿的依赖范围未知；请使用明确的已有缓存值快照导出。");
        RequireExtension(path, ".xlsx");
        GuardOriginal(path, snapshot, discardUnsupported);
        if (!discardUnsupported && snapshot.FileState?.CompatibilityIssues.Length > 0)
            throw new WorkbookCompatibilityException(snapshot.FileState.CompatibilityIssues);
        var index = PortableIndex(snapshot);
        var sheets = snapshot.Sheets.Select(sheet => sheet.WithCells(sheet.Cells.ToImmutableDictionary(p => p.Key, p =>
        {
            var cell = p.Value;
            if (cell.SpillOwner is { } owner && index.NonPortableCells.Contains(owner))
                return NormalizeHostError(new CellSnapshot(CellContent.FromLiteral(cell.Value, cell.Content.StyleId)));
            if (cell.Content.Kind != CellContentKind.Formula)
                return NormalizeHostError(cell);
            var address = new CellAddress(sheet.Id, p.Key.Row, p.Key.Column);
            if (index.NonPortableCells.Contains(address))
            {
                var result = cell.Calculation;
                if (result?.Value is null) throw new InvalidOperationException("脚本或其传递依赖缺少结果，不能导出可移植公式。");
                if (!allowStaleCaches && (result.InputRevision != snapshot.Revision || result.State is not (CalculationState.Fresh or CalculationState.Error)))
                    throw new InvalidOperationException("需要冻结的结果可能过期；请重算或明确选择已有缓存。");
                return NormalizeHostError(new CellSnapshot(CellContent.FromLiteral(result.Value.Value, cell.Content.StyleId)));
            }
            // Expand names so no remaining definition can hide a UDF. Stale standard caches are omitted;
            // Excel will recalculate retained formulas via fullCalcOnLoad.
            var formula = index.Nodes[address].PortableFormula!;
            var cache = cell.Calculation is { } c && c.InputRevision == snapshot.Revision
                && c.State is CalculationState.Fresh or CalculationState.Error ? c : null;
            return new CellSnapshot(CellContent.FromFormula(formula, cell.Content.StyleId, cell.Content.DynamicArray), cache);
        }))).ToArray();
        var exported = snapshot.WithDefinedNames([]).WithSheets(sheets);
        var bytes = TypedWorkbookXml.Write(exported, false, CancellationToken.None);
        AtomicWrite(path, bytes, temp =>
        {
            var read = TypedWorkbookXml.Read(ReadFile(temp, new(CancellationToken.None)), new(CancellationToken.None));
            if (!ContentEquals(exported, read.Snapshot, allowRowHeightRounding: true)) throw new InvalidDataException("可移植公式导出后校验未通过。");
        }, CancellationToken.None);
    }

    // Export must follow derived-cell dependencies too: freezing an array also freezes its
    // children and consumers, including formulas that reference a child rather than the anchor.
    private static WorkbookDependencyIndex PortableIndex(WorkbookSnapshot snapshot)
    {
        var spills = new Dictionary<CellAddress, FormulaSpill>();
        foreach (var sheet in snapshot.Sheets)
        foreach (var (position, cell) in sheet.Cells)
            if (cell.Content.DynamicArray && cell.Calculation?.Array is { } array)
            {
                var address = new CellAddress(sheet.Id, position.Row, position.Column);
                spills.Add(address, new(address, array));
            }
        return new(snapshot, () => { }, spills);
    }

    private static CellSnapshot NormalizeHostError(CellSnapshot cell) =>
        cell.Value is { Kind: CellValueKind.Error, ErrorCode: CellErrorCode.Cycle or CellErrorCode.Script }
            ? new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Value), cell.Content.StyleId)) : cell;
}
