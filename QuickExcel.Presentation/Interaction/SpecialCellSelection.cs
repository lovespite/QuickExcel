using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public enum SpecialCellKind { Blank, Constant, Formula, Error, Visible }
public sealed record SpecialCellSelectionResult(IReadOnlyList<CellRectangle> Ranges, int Cells);

/// <summary>Plans a complete selection without mutating navigation or allocating the used rectangle for sparse queries.</summary>
public static class SpecialCellSelection
{
    public const int MaxCells = 1_000_000;
    public const int MaxWork = 2_000_000;
    public static SpecialCellSelectionResult Find(WorkbookSnapshot book, SheetId id, IEnumerable<CellRectangle> selection,
        SpecialCellKind kind, CancellationToken cancellationToken = default, int workBudget = MaxWork)
    {
        if (!Enum.IsDefined(kind) || workBudget is < 1 or > MaxWork) throw new ArgumentOutOfRangeException(nameof(kind));
        var sheet = book.GetSheet(id); var ranges = selection.Take(65).ToArray();
        if (ranges.Length > 64) throw new ArgumentException(MessageText.Get("定位条件最多接受 64 个选区。"));
        foreach (var r in ranges)
        {
            WorkbookLimits.ValidatePosition(r.Top, r.Left); WorkbookLimits.ValidatePosition(r.Bottom, r.Right);
            if (r.Rows < 1 || r.Columns < 1) throw new ArgumentException(MessageText.Get("选区端点无效。"));
        }
        var work = 0;
        void Visit()
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (++work > workBudget) throw new InvalidOperationException(MessageText.Get("定位条件达到扫描上限，请缩小范围。"));
        }
        bool InScope(CellRectangle rect)
        {
            foreach (var r in ranges) { Visit(); if (r.Contains(rect)) return true; }
            return false;
        }
        var merges = sheet.MergedCells.Select(CellRectangle.Parse).ToArray();
        var hidden = kind == SpecialCellKind.Visible
            ? WorksheetFilterEngine.Evaluate(book, id, WorksheetFilterEngine.CanUseSourceCaches(book), cancellationToken) : null;
        bool Allowed(int row, int col) => ProtectionSafety.CanSelect(book, sheet, row, col) &&
            (kind != SpecialCellKind.Visible || (sheet.Rows.GetValueOrDefault(row)?.Hidden != true && sheet.Columns.GetValueOrDefault(col)?.Hidden != true && !hidden!.Contains(row)));
        var selected = new HashSet<CellPosition>();
        void Match(CellPosition position)
        {
            Visit();
            var target = new CellRectangle(position.Row, position.Column, position.Row, position.Column);
            if (!InScope(target)) return;
            foreach (var merge in merges)
            {
                Visit(); if (!merge.Contains(position.Row, position.Column)) continue;
                if (merge.Top != position.Row || merge.Left != position.Column || !InScope(merge)) return;
                target = merge; break;
            }
            var cell = sheet.GetCell(position.Row, position.Column);
            var formula = cell?.Content.Kind == CellContentKind.Formula; var spill = cell?.SpillOwner is not null;
            if (kind == SpecialCellKind.Error && (formula || spill) &&
                (cell!.Calculation?.Value is null || cell.Calculation.InputRevision != book.Revision || cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)))
                throw new InvalidOperationException(MessageText.Get("错误值定位需要最新计算结果，请先计算。"));
            var matches = kind switch
            {
                SpecialCellKind.Blank => !formula && !spill && (cell is null || cell.Content.LiteralValue.Kind == CellValueKind.Blank),
                SpecialCellKind.Constant => !formula && !spill && cell is not null && cell.Content.LiteralValue.Kind != CellValueKind.Blank,
                SpecialCellKind.Formula => formula,
                SpecialCellKind.Error => cell?.Value.Kind == CellValueKind.Error,
                _ => true
            };
            if (!matches) return;
            for (var row = target.Top; row <= target.Bottom; row++)
                for (var col = target.Left; col <= target.Right; col++) { Visit(); if (!Allowed(row, col)) return; }
            for (var row = target.Top; row <= target.Bottom; row++)
                for (var col = target.Left; col <= target.Right; col++)
                {
                    selected.Add(new(row, col));
                    if (selected.Count > MaxCells) throw new InvalidOperationException(MessageText.Get("定位条件达到扫描上限，请缩小范围。"));
                }
        }
        if (kind is SpecialCellKind.Blank or SpecialCellKind.Visible)
        {
            if (ranges.Sum(r => r.Count) > MaxCells) throw new InvalidOperationException(MessageText.Get("定位条件达到扫描上限，请缩小范围。"));
            foreach (var r in ranges)
                for (var row = r.Top; row <= r.Bottom; row++)
                    for (var col = r.Left; col <= r.Right; col++) Match(new(row, col));
        }
        else foreach (var position in sheet.Cells.Keys) Match(position);
        cancellationToken.ThrowIfCancellationRequested();
        var output = new List<CellRectangle>(); var active = new Dictionary<(int Left, int Right), int>(); var previousRow = -2;
        foreach (var group in selected.OrderBy(p => p.Row).ThenBy(p => p.Column).GroupBy(p => p.Row))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (group.Key != previousRow + 1) active.Clear();
            var next = new Dictionary<(int, int), int>(); var columns = group.Select(p => p.Column).ToArray();
            for (var i = 0; i < columns.Length; i++)
            {
                var left = columns[i]; var right = left;
                while (i + 1 < columns.Length && columns[i + 1] == right + 1) right = columns[++i];
                var key = (left, right);
                if (active.TryGetValue(key, out var index)) output[index] = output[index] with { Bottom = group.Key };
                else { index = output.Count; output.Add(new(group.Key, left, group.Key, right)); }
                next.Add(key, index);
            }
            if (output.Count > SelectionStatistics.MaxRanges) throw new InvalidOperationException(MessageText.Get("匹配结果超过 64 个连续区域，请缩小定位范围。"));
            active = next; previousRow = group.Key;
        }
        return new(output, selected.Count);
    }
}
