using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record BorderStroke(int Style = 1, string? Color = null);

/// <summary>Null leaves an edge unchanged; style zero removes it. Diagonals share one XLSX stroke.</summary>
public sealed record CellBorderPatch(BorderStroke? Left = null, BorderStroke? Right = null,
    BorderStroke? Top = null, BorderStroke? Bottom = null, BorderStroke? InsideHorizontal = null,
    BorderStroke? InsideVertical = null, BorderStroke? Diagonal = null, bool? DiagonalUp = null, bool? DiagonalDown = null);

public enum BorderPreset { Bottom, Top, Left, Right, None, All, Outside, ThickOutside, DoubleBottom, ThickBottom, TopBottom, TopThickBottom, TopDoubleBottom, Inside, InsideHorizontal, InsideVertical }

public static class CellBorders
{
    public static CellBorderPatch Preset(BorderPreset preset, int style = 1, string? color = null)
    {
        var line = new BorderStroke(style, color); var thick = new BorderStroke(5, color); var doubleLine = new BorderStroke(6, color);
        return preset switch
        {
            BorderPreset.Bottom => new(Bottom: line), BorderPreset.Top => new(Top: line),
            BorderPreset.Left => new(Left: line), BorderPreset.Right => new(Right: line),
            BorderPreset.None => new(new(0), new(0), new(0), new(0), new(0), new(0), new(0), false, false),
            BorderPreset.All => new(line, line, line, line, line, line),
            BorderPreset.Outside => new(line, line, line, line), BorderPreset.ThickOutside => new(thick, thick, thick, thick),
            BorderPreset.DoubleBottom => new(Bottom: doubleLine), BorderPreset.ThickBottom => new(Bottom: thick),
            BorderPreset.TopBottom => new(Top: line, Bottom: line), BorderPreset.TopThickBottom => new(Top: line, Bottom: thick),
            BorderPreset.TopDoubleBottom => new(Top: line, Bottom: doubleLine),
            BorderPreset.Inside => new(InsideHorizontal: line, InsideVertical: line),
            BorderPreset.InsideHorizontal => new(InsideHorizontal: line), BorderPreset.InsideVertical => new(InsideVertical: line),
            _ => throw new ArgumentOutOfRangeException(nameof(preset))
        };
    }

    public static int Weight(int style) => style switch { 0 => 0, 7 => 1, 2 or 8 or 10 or 12 or 13 => 3, 6 => 4, 5 => 5, _ => 2 };
}

/// <summary>One atomic edit over the visible selected cells. Shared edges also clear opposing direct borders.</summary>
public sealed class SetCellBordersCommand : DocumentCommand
{
    private readonly CellAddress[] _addresses;
    private readonly CellBorderPatch _patch;
    public SetCellBordersCommand(IEnumerable<CellAddress> addresses, CellBorderPatch patch)
    {
        _addresses = addresses.Distinct().Take(20_001).ToArray(); _patch = patch;
        if (_addresses.Length > 20_000) throw new ArgumentException("边框设置上限为 20,000 格。");
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        foreach (var stroke in new[] { _patch.Left, _patch.Right, _patch.Top, _patch.Bottom, _patch.InsideHorizontal, _patch.InsideVertical, _patch.Diagonal }.OfType<BorderStroke>())
            if (stroke.Style is < 0 or > 13 || stroke.Color is { } color && (color.Length != 8 || color.Any(c => !Uri.IsHexDigit(c))))
                throw new ArgumentException("边框线型或颜色无效。");
        var selected = _addresses.ToHashSet();
        var pending = new Dictionary<CellAddress, StyleDefinition>();
        foreach (var group in _addresses.GroupBy(a => a.SheetId))
        {
            var sheet = snapshot.GetSheet(group.Key);
            var merged = new Dictionary<CellPosition, CellRectangle>();
            foreach (var merge in sheet.MergedCells.Select(CellRectangle.Parse))
            {
                if (!group.Any(a => merge.Contains(a.Row, a.Column))) continue;
                if (merge.Count > selected.Count || Enumerable.Range(merge.Top, merge.Bottom - merge.Top + 1).Any(r =>
                    Enumerable.Range(merge.Left, merge.Right - merge.Left + 1).Any(c => !selected.Contains(new(sheet.Id, r, c)))))
                    throw new InvalidOperationException("请选中完整合并区域后设置边框。");
                for (var r = merge.Top; r <= merge.Bottom; r++) for (var c = merge.Left; c <= merge.Right; c++) merged[new(r, c)] = merge;
            }
            StyleDefinition Get(CellAddress a) => pending.GetValueOrDefault(a) ?? snapshot.Styles[sheet.GetStyleId(a.Row, a.Column)];
            void Edge(CellAddress a, int side, BorderStroke stroke)
            {
                var s = Get(a); var color = stroke.Color; // null explicitly means automatic, not the previous color.
                pending[a] = side switch
                {
                    0 => s with { LeftBorder = stroke.Style, LeftBorderColor = color },
                    1 => s with { RightBorder = stroke.Style, RightBorderColor = color },
                    2 => s with { TopBorder = stroke.Style, TopBorderColor = color },
                    _ => s with { BottomBorder = stroke.Style, BottomBorderColor = color }
                };
            }
            foreach (var a in group)
            {
                WorkbookLimits.ValidatePosition(a.Row, a.Column);
                var hasMerge = merged.TryGetValue(new(a.Row, a.Column), out var merge);
                foreach (var (side, dr, dc, outside, inside) in new[] {
                    (0, 0, -1, _patch.Left, _patch.InsideVertical), (1, 0, 1, _patch.Right, _patch.InsideVertical),
                    (2, -1, 0, _patch.Top, _patch.InsideHorizontal), (3, 1, 0, _patch.Bottom, _patch.InsideHorizontal) })
                {
                    var nr = a.Row + dr; var nc = a.Column + dc;
                    if (hasMerge && merge.Contains(nr, nc)) continue;
                    var next = new CellAddress(a.SheetId, nr, nc); var within = selected.Contains(next);
                    if ((within ? inside : outside) is not { } stroke) continue;
                    Edge(a, side, stroke);
                    if (nr >= 0 && nc >= 0 && nr < WorkbookLimits.MaxRows && nc < WorkbookLimits.MaxColumns)
                    {
                        // Both sides of an internal edge agree; outside neighbors only lose the conflicting edge.
                        Edge(next, side ^ 1, within ? stroke : new(0));
                    }
                }
                if ((!hasMerge || a.Row == merge.Top && a.Column == merge.Left) && (_patch.Diagonal is not null || _patch.DiagonalUp.HasValue || _patch.DiagonalDown.HasValue))
                {
                    var s = Get(a);
                    pending[a] = s with { DiagonalBorder = _patch.Diagonal?.Style ?? s.DiagonalBorder,
                        DiagonalBorderColor = _patch.Diagonal is null ? s.DiagonalBorderColor : _patch.Diagonal.Color,
                        DiagonalUp = _patch.DiagonalUp ?? s.DiagonalUp, DiagonalDown = _patch.DiagonalDown ?? s.DiagonalDown };
                }
            }
        }
        var styles = snapshot.Styles.ToList(); var ids = new Dictionary<StyleDefinition, int>();
        for (var i = 0; i < styles.Count; i++) ids.TryAdd(styles[i], i);
        var edits = new List<CellEdit>();
        foreach (var (a, style) in pending)
        {
            var sheet = snapshot.GetSheet(a.SheetId);
            if (style == snapshot.Styles[sheet.GetStyleId(a.Row, a.Column)]) continue;
            if (!ids.TryGetValue(style, out var id))
            {
                if (styles.Count >= WorkbookLimits.MaxStyles) throw new InvalidOperationException("样式数量超限，边框未应用。");
                id = styles.Count; ids.Add(style, id); styles.Add(style);
            }
            edits.Add(new(a, (sheet.GetCell(a.Row, a.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(id)));
        }
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets,
            snapshot.ScriptSource, styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles);
        var draft = new SetCellsCommand(edits).Prepare(styled);
        return draft.Changes.IsEmpty ? new(snapshot, [], false) : draft with { CalculationUnchanged = true };
    }
}
