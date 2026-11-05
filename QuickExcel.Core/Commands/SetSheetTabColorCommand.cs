using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Changes only the tab appearance; null restores the default appearance.</summary>
public sealed class SetSheetTabColorCommand : DocumentCommand
{
    public SetSheetTabColorCommand(SheetId sheetId, string? color)
    {
        SheetId = sheetId;
        if (color is null) return;
        var value = color.Trim();
        if (value.StartsWith('#')) value = value[1..];
        if (value.Length is not (6 or 8) || value.Any(c => !Uri.IsHexDigit(c)))
            throw new ArgumentException(MessageText.Get("请输入有效的十六进制颜色（#RRGGBB 或 AARRGGBB）。"), nameof(color));
        Color = (value.Length == 6 ? "FF" + value : value).ToUpperInvariant();
    }

    public SheetId SheetId { get; }
    public string? Color { get; }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(SheetId);
        if (string.Equals(sheet.TabColor, Color, StringComparison.OrdinalIgnoreCase)) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == SheetId ? s.WithTabColor(Color) : s)),
            [], false, CalculationUnchanged: true);
    }
}
