using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed class RenameSheetCommand(SheetId sheetId, string name) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var old = snapshot.GetSheet(sheetId);
        WorkbookLimits.ValidateSheetName(name);
        if (old.Name == name) return new(snapshot, [], false);
        var renamed = StructureEdits.Transform(snapshot, (node, _) => node switch
        {
            ReferenceNode r when string.Equals(r.Sheet, old.Name, StringComparison.OrdinalIgnoreCase) => r with { Sheet = name },
            NameNode n when string.Equals(n.Sheet, old.Name, StringComparison.OrdinalIgnoreCase) => n with { Sheet = name },
            _ => node
        }, (sheet, cells) => sheet.Id == sheetId ? sheet.WithCells(cells).WithName(name) : sheet.WithCells(cells));
        return StructureEdits.Draft(snapshot, renamed);
    }
}

public sealed class DefineNameCommand(DefinedName definition) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var names = snapshot.DefinedNames.Where(n => n.Scope != definition.Scope || !n.Name.Equals(definition.Name, StringComparison.OrdinalIgnoreCase)).Append(definition);
        return StructureEdits.Draft(snapshot, snapshot.WithDefinedNames(names));
    }
}

public enum SheetAxis { Row, Column }

/// <summary>Insertion/deletion uses cut-style reference semantics: dollar flags remain, targets follow structure.</summary>
public sealed class ShiftAxisCommand(SheetId sheetId, SheetAxis axis, int index, int count, bool delete = false) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (!Enum.IsDefined(axis) || index < 0 || count < 1 || index > limit - count) throw new ArgumentOutOfRangeException(nameof(index));
        var target = snapshot.GetSheet(sheetId);
        if (delete && (!target.Printing.RowBreaks.IsEmpty || !target.Printing.ColumnBreaks.IsEmpty))
            throw new NotSupportedException("含分页信息的工作表暂不支持删除行列；Excel 的分页重建规则尚未完整对齐，操作未应用。");
        int? Position(int value)
        {
            if (!delete) return value < index ? value : checked(value + count);
            return value < index ? value : value < index + count ? null : value - count;
        }
        (int, int)? Interval(int start, int end)
        {
            if (!delete) return (start >= index ? start + count : start, end >= index ? end + count : end);
            var left = start < index ? start : Math.Max(start, index + count);
            var right = end >= index + count ? end : Math.Min(end, index - 1);
            return left > right ? null : (Position(left)!.Value, Position(right)!.Value);
        }
        var transformed = StructureEdits.Axis(snapshot, target, axis, Position, Interval);
        var view = target.View;
        var boundary = axis == SheetAxis.Row ? view.FrozenRows : view.FrozenColumns;
        if (index < boundary)
        {
            var nextBoundary = delete ? boundary - Math.Min(count, boundary - index) : checked(boundary + count);
            var nextView = axis == SheetAxis.Row ? view with { FrozenRows = nextBoundary, PaneTopLeftCell = view.PaneTopLeftCell with { Row = nextBoundary } }
                : view with { FrozenColumns = nextBoundary, PaneTopLeftCell = view.PaneTopLeftCell with { Column = nextBoundary } };
            if (!nextView.HasFreeze)
                nextView = nextView with { FrozenSplit = false, PaneTopLeftCell = default, ActivePane = SheetPane.TopLeft,
                    TopLeft = view.Selection, TopRight = null, BottomLeft = null, BottomRight = null };
            else if (nextBoundary == 0)
            {
                // Removing one frozen axis collapses its panes. Keep the active selection in its surviving pane.
                var activeSelection = view.Selection;
                nextView = nextView.WithActiveCell(view.ActiveCell);
                nextView = axis == SheetAxis.Row ? nextView with { BottomLeft = null, BottomRight = null }
                    : nextView with { TopRight = null, BottomRight = null };
                if (activeSelection is not null)
                    nextView = nextView.ActivePane switch
                    {
                        SheetPane.TopRight => nextView with { TopRight = activeSelection },
                        SheetPane.BottomLeft => nextView with { BottomLeft = activeSelection },
                        _ => nextView with { TopLeft = activeSelection }
                    };
            }
            transformed = transformed.WithSheets(transformed.Sheets.Select(s => s.Id == sheetId ? s.WithView(nextView) : s));
        }
        return StructureEdits.Draft(snapshot, transformed);
    }
}

/// <summary>Moves a contiguous row/column block to its final index after removal. Splitting a reference range is rejected.</summary>
public sealed class MoveAxisCommand(SheetId sheetId, SheetAxis axis, int start, int count, int destination) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (!Enum.IsDefined(axis) || start < 0 || destination < 0 || count < 1 || start > limit - count || destination > limit - count)
            throw new ArgumentOutOfRangeException(nameof(start));
        if (start == destination) return new(snapshot, [], false);
        var target = snapshot.GetSheet(sheetId);
        if (!target.Printing.RowBreaks.IsEmpty || !target.Printing.ColumnBreaks.IsEmpty)
            throw new NotSupportedException("含分页信息的工作表暂不支持移动行列；操作未应用。");
        int? Position(int value)
        {
            if (value >= start && value < start + count) return destination + value - start;
            var removed = value < start ? value : value - count;
            return removed >= destination ? removed + count : removed;
        }
        (int, int)? Interval(int first, int last)
        {
            var cuts = new[] { first, last + 1, start, start + count, destination, destination + count }
                .Where(v => v >= first && v <= last + 1).Distinct().Order().ToArray();
            var segments = new List<(int Start, int End)>();
            for (var i = 0; i < cuts.Length - 1; i++) segments.Add((Position(cuts[i])!.Value, Position(cuts[i + 1] - 1)!.Value));
            var min = segments.Min(s => s.Start); var max = segments.Max(s => s.End);
            if (max - min != last - first) throw new NotSupportedException("此移动会把引用区域拆成不连续区域，操作未应用。");
            return (min, max);
        }
        return StructureEdits.Draft(snapshot, StructureEdits.Axis(snapshot, target, axis, Position, Interval));
    }
}

public sealed class MoveCellsCommand(SheetId sourceSheet, CellPosition first, CellPosition last,
    SheetId destinationSheet, CellPosition destination) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookLimits.ValidatePosition(first.Row, first.Column); WorkbookLimits.ValidatePosition(last.Row, last.Column);
        if (first.Row > last.Row || first.Column > last.Column) throw new ArgumentException("移动区域顺序无效。");
        WorkbookLimits.ValidatePosition(destination.Row, destination.Column);
        var dr = destination.Row - first.Row; var dc = destination.Column - first.Column;
        WorkbookLimits.ValidatePosition(last.Row + dr, last.Column + dc);
        if (sourceSheet == destinationSheet && dr == 0 && dc == 0) return new(snapshot, [], false);
        var source = snapshot.GetSheet(sourceSheet); var target = snapshot.GetSheet(destinationSheet);
        if (source.Printing.HasReferences || target.Printing.HasReferences)
            throw new NotSupportedException("含打印区域或重复标题的工作表暂不支持移动单元格，请先使用复制；操作未应用。");
        StructureEdits.RequireSimpleStructure(source); StructureEdits.RequireSimpleStructure(target);
        bool Inside(CellPosition p) => p.Row >= first.Row && p.Row <= last.Row && p.Column >= first.Column && p.Column <= last.Column;
        var sourceRange = new ReferenceNode(source.Name, first.Row, first.Column, last.Row, last.Column, IsRange: true);
        FormulaNode Transform(FormulaNode node, SheetId context)
        {
            if (node is not ReferenceNode r) return node;
            var name = r.Sheet ?? snapshot.GetSheet(context).Name;
            if (!name.Equals(source.Name, StringComparison.OrdinalIgnoreCase)) return node;
            var intersects = r.Row <= last.Row && r.EndRow >= first.Row && r.Column <= last.Column && r.EndColumn >= first.Column;
            if (!intersects) return node;
            if (!sourceRange.Contains(r.Row, r.Column) || !sourceRange.Contains(r.EndRow, r.EndColumn))
                throw new NotSupportedException("移动仅覆盖了引用区域的一部分，无法保留连续区域语义，操作未应用。");
            return r with { Sheet = target.Name, Row = r.Row + dr, EndRow = r.EndRow + dr, Column = r.Column + dc, EndColumn = r.EndColumn + dc };
        }
        var transformed = StructureEdits.Transform(snapshot, Transform);
        var moved = transformed.GetSheet(sourceSheet).Cells.Where(p => Inside(p.Key)).ToArray();
        var sheets = transformed.Sheets.Select(sheet =>
        {
            var cells = sheet.Cells.ToBuilder();
            if (sheet.Id == sourceSheet) foreach (var p in sheet.Cells.Keys.Where(Inside)) cells.Remove(p);
            if (sheet.Id == destinationSheet)
            {
                foreach (var p in sheet.Cells.Keys.Where(p => p.Row >= destination.Row && p.Row <= last.Row + dr && p.Column >= destination.Column && p.Column <= last.Column + dc)) cells.Remove(p);
                foreach (var (p, original) in moved)
                {
                    var cell = original;
                    if (cell.Content.Kind == CellContentKind.Formula && sourceSheet != destinationSheet)
                    {
                        var expression = FormulaSyntax.Map(FormulaSyntax.Parse(cell.Content.FormulaText), n => n switch
                        {
                            ReferenceNode r when r.Sheet is null => r with { Sheet = source.Name },
                            NameNode n1 when n1.Sheet is null => n1 with { Sheet = source.Name },
                            _ => n
                        });
                        cell = new(CellContent.FromFormula("=" + FormulaSyntax.Print(expression), cell.Content.StyleId));
                    }
                    cells[new(p.Row + dr, p.Column + dc)] = cell;
                }
            }
            return sheet.WithCells(cells.ToImmutable());
        });
        return StructureEdits.Draft(snapshot, transformed.WithSheets(sheets));
    }
}

internal static class StructureEdits
{
    internal static void RequireSimpleStructure(WorksheetSnapshot sheet)
    {
        if (!sheet.Images.IsEmpty || !sheet.MergedCells.IsEmpty)
            throw new NotSupportedException("带合并单元格或图片的工作表暂不支持结构移动；操作未应用。");
    }

    internal static WorkbookSnapshot Axis(WorkbookSnapshot snapshot, WorksheetSnapshot target, SheetAxis axis,
        Func<int, int?> position, Func<int, int, (int Start, int End)?> interval)
    {
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        FormulaNode Transform(FormulaNode node, SheetId context)
        {
            if (node is not ReferenceNode r || !string.Equals(r.Sheet ?? snapshot.GetSheet(context).Name, target.Name, StringComparison.OrdinalIgnoreCase)
                || axis == SheetAxis.Row && r.WholeColumns) return node;
            var mapped = interval(axis == SheetAxis.Row ? r.Row : r.Column, axis == SheetAxis.Row ? r.EndRow : r.EndColumn);
            if (mapped is null) return new ValueNode(CellValue.FromError(CellErrorCode.Reference));
            if (mapped.Value.End >= limit) throw new ArgumentException("结构编辑会使公式引用越界，操作未应用。");
            return axis == SheetAxis.Row ? r with { Row = mapped.Value.Start, EndRow = mapped.Value.End }
                : r with { Column = mapped.Value.Start, EndColumn = mapped.Value.End };
        }
        return StructureEdits.Transform(snapshot, Transform, (sheet, cells) =>
        {
            if (sheet.Id != target.Id) return sheet.WithCells(cells);
            var relocated = ImmutableDictionary.CreateBuilder<CellPosition, CellSnapshot>();
            foreach (var (p, cell) in cells)
            {
                var next = position(axis == SheetAxis.Row ? p.Row : p.Column);
                if (next is null) continue;
                if (next >= limit) throw new ArgumentException("结构编辑会使单元格越界，操作未应用。");
                relocated.Add(axis == SheetAxis.Row ? p with { Row = next.Value } : p with { Column = next.Value }, cell);
            }
            IEnumerable<KeyValuePair<int, T>> Metadata<T>(IEnumerable<KeyValuePair<int, T>> entries)
            {
                foreach (var (key, value) in entries)
                {
                    var next = position(key);
                    if (next is null) continue;
                    if (next >= limit) throw new ArgumentException("结构编辑会使行列属性越界，操作未应用。");
                    yield return new(next.Value, value);
                }
            }
            CellRectangle? MapRectangle(CellRectangle rectangle)
            {
                var mapped = interval(axis == SheetAxis.Row ? rectangle.Top : rectangle.Left, axis == SheetAxis.Row ? rectangle.Bottom : rectangle.Right);
                if (mapped is null) return null;
                if (mapped.Value.Start < 0 || mapped.Value.End >= limit) throw new ArgumentException("结构编辑会使图片或合并区域越界，操作未应用。");
                return axis == SheetAxis.Row ? rectangle with { Top = mapped.Value.Start, Bottom = mapped.Value.End }
                    : rectangle with { Left = mapped.Value.Start, Right = mapped.Value.End };
            }
            var merges = new List<CellRectangle>();
            foreach (var text in sheet.MergedCells)
            {
                var original = CellRectangle.Parse(text);
                if (MapRectangle(original) is not { } mapped) continue;
                // A partial move must retain the top-left owner, otherwise its value would be hidden.
                var owner = position(axis == SheetAxis.Row ? original.Top : original.Left);
                if (owner is null || owner != (axis == SheetAxis.Row ? mapped.Top : mapped.Left))
                    throw new NotSupportedException("操作会移除或移动合并区域的主单元格，请先取消合并。");
                if (merges.Any(m => m.Top <= mapped.Bottom && m.Bottom >= mapped.Top && m.Left <= mapped.Right && m.Right >= mapped.Left))
                    throw new NotSupportedException("操作会使合并区域重叠。");
                merges.Add(mapped);
            }
            var images = new List<SheetImage>();
            foreach (var image in sheet.Images)
            {
                var original = new CellRectangle(image.RowIndex, image.ColumnIndex, image.RowIndex + image.RowSpan - 1, image.ColumnIndex + image.ColumnSpan - 1);
                if (MapRectangle(original) is not { } mapped) continue;
                images.Add(new(image.Data, image.Extension, mapped.Top, mapped.Left, mapped.Bottom - mapped.Top + 1, mapped.Right - mapped.Left + 1, image.PlaceInCell));
            }
            return new(sheet.Id, sheet.Name, relocated, axis == SheetAxis.Row ? Metadata(sheet.Rows) : sheet.Rows,
                axis == SheetAxis.Column ? Metadata(sheet.Columns) : sheet.Columns, merges.Select(m => m.ToString()), images,
                sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility,
                PrintingStructure.Axis(sheet.Printing, axis, position, interval));
        });
    }

    internal static WorkbookSnapshot Transform(WorkbookSnapshot snapshot, Func<FormulaNode, SheetId, FormulaNode> transform,
        Func<WorksheetSnapshot, ImmutableDictionary<CellPosition, CellSnapshot>, WorksheetSnapshot>? finish = null)
    {
        string Formula(string text, SheetId context)
        {
            try
            {
                var changed = false;
                var expression = FormulaSyntax.Map(FormulaSyntax.Parse(text), node =>
                {
                    var next = transform(node, context);
                    changed |= next != node;
                    return next;
                });
                return changed ? "=" + FormulaSyntax.Print(expression) : text;
            }
            catch (FormatException e) { throw new NotSupportedException("存在无法安全变换的公式，结构操作未应用：" + e.Message, e); }
        }
        var sheets = snapshot.Sheets.Select(sheet =>
        {
            var cells = sheet.Cells.ToBuilder();
            foreach (var (p, cell) in sheet.Cells)
            {
                if (cell.Content.Kind != CellContentKind.Formula) continue;
                var formula = Formula(cell.Content.FormulaText, sheet.Id);
                if (formula != cell.Content.FormulaText) cells[p] = new(CellContent.FromFormula(formula, cell.Content.StyleId));
            }
            return finish?.Invoke(sheet, cells.ToImmutable()) ?? sheet.WithCells(cells.ToImmutable());
        }).ToArray();
        var names = snapshot.DefinedNames.Select(n => n with { Formula = Formula(n.Formula, n.Scope ?? snapshot.Sheets[0].Id) });
        // Rename and names must validate together.
        return new(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, sheets, snapshot.ScriptSource,
            snapshot.Styles, snapshot.ScriptModules, snapshot.FileState, names, snapshot.ActiveSheetId, snapshot.Window);
    }

    internal static CommandDraft Draft(WorkbookSnapshot before, WorkbookSnapshot after)
    {
        var changes = new List<CellChange>();
        foreach (var sheet in after.Sheets)
        {
            var old = before.GetSheet(sheet.Id);
            foreach (var p in old.Cells.Keys.Union(sheet.Cells.Keys))
            {
                var a = old.Cells.GetValueOrDefault(p); var b = sheet.Cells.GetValueOrDefault(p);
                if (a != b) changes.Add(new(new(sheet.Id, p.Row, p.Column), a, b));
            }
        }
        return new(after, changes.ToImmutableArray(), true);
    }
}
