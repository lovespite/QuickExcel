using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed class AdjustDecimalPlacesCommand : DocumentCommand
{
    private readonly ImmutableArray<CellAddress> _addresses;
    private readonly int _delta;
    public AdjustDecimalPlacesCommand(IEnumerable<CellAddress> addresses, int delta)
    {
        if (delta is not (1 or -1)) throw new ArgumentOutOfRangeException(nameof(delta));
        _addresses = addresses.Take(20_001).ToImmutableArray(); _delta = delta;
        if (_addresses.Length > 20_000) throw new ArgumentException("小数位编辑上限为 20,000 格。");
    }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var styles = snapshot.Styles.ToList(); var ids = new Dictionary<StyleDefinition, int>();
        for (var i = 0; i < styles.Count; i++) ids.TryAdd(styles[i], i);
        var edits = new List<CellEdit>();
        foreach (var address in _addresses.Distinct())
        {
            var sheet = snapshot.GetSheet(address.SheetId); var cell = sheet.GetCell(address.Row, address.Column);
            var style = styles[sheet.GetStyleId(address.Row, address.Column)];
            if (!CellDisplayFormatter.TryAdjustDecimalPlaces(style, cell?.Value ?? CellValue.Blank, _delta, out var adjusted))
                throw new InvalidOperationException("选区含日期、时间、文本、分数或未支持的自定义格式，请先选择数值格式；本次未修改任何单元格。");
            if (adjusted == style) continue;
            if (!ids.TryGetValue(adjusted, out var id))
            {
                if (styles.Count >= 4096) throw new InvalidOperationException("样式数量超过 4096，操作未应用。");
                id = styles.Count; styles.Add(adjusted); ids.Add(adjusted, id);
            }
            edits.Add(new(address, (cell?.Content ?? CellContent.FromLiteral(CellValue.Blank)).WithStyle(id)));
        }
        if (edits.Count == 0) return new(snapshot, [], false);
        var styled = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets,
            snapshot.ScriptSource, styles, snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles);
        var draft = new SetCellsCommand(edits).Prepare(styled);
        return draft with { CalculationUnchanged = true };
    }
}
