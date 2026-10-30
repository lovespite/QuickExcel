using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record ValidationFailure(CellAddress Address, ValidationAlert Alert, string Title, string Message);
public sealed class DataValidationException(IReadOnlyList<ValidationFailure> failures) : InvalidOperationException(
    string.Join("\n", failures.Take(5).Select(f => $"{new CellRectangle(f.Address.Row, f.Address.Column, f.Address.Row, f.Address.Column)}：{f.Title} {f.Message}")) +
    (failures.Count > 5 ? $"\n另有 {failures.Count - 5} 个不符合规则的单元格。" : "") + "\n整批修改尚未应用。")
{
    public IReadOnlyList<ValidationFailure> Failures { get; } = failures;
    public bool CanOverride => Failures.All(f => f.Alert != ValidationAlert.Stop);
}

public interface IWorkbookValidation
{
    IReadOnlyList<ValidationFailure> Check(WorkbookSnapshot snapshot, IReadOnlyList<CellAddress> addresses);
    IReadOnlyList<CellValue> ListValues(WorkbookSnapshot snapshot, SheetId sheet, CellPosition position, WorksheetValidation rule);
}

public sealed class SetValidationCommand(SheetId sheetId, WorksheetValidation rule) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        rule.ValidateBinding(snapshot, sheetId);
        var same = sheet.Data.Validations.FirstOrDefault(v => v.Ranges.SequenceEqual(rule.Ranges));
        if (same?.ContentEquals(rule) == true) return new(snapshot, [], false);
        if (sheet.Data.Validations.Any(v => v != same && v.Ranges.Any(a => rule.Ranges.Any(b => ValidationSafety.Overlaps(a, b)))))
            throw new InvalidOperationException("新规则与现有验证重叠，请先清除原规则；未修改工作簿。");
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ?
            s.WithData(s.Data with { Validations = same is null ? s.Data.Validations.Add(rule) : s.Data.Validations.Replace(same, rule) }) : s)), [], false) { CalculationUnchanged = true };
    }
}

/// <summary>Remove whole intersecting rules; callers must show their complete ranges before confirmation.</summary>
public sealed class RemoveValidationCommand(SheetId sheetId, CellRectangle range) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId); var remaining = sheet.Data.Validations.Where(v => !v.Ranges.Any(r => ValidationSafety.Overlaps(r, range))).ToImmutableArray();
        if (remaining.Length == sheet.Data.Validations.Length) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(s.Data with { Validations = remaining }) : s)), [], false) { CalculationUnchanged = true };
    }
}

internal static class ValidationSafety
{
    internal static bool Overlaps(CellRectangle a, CellRectangle b) => a.Top <= b.Bottom && b.Top <= a.Bottom && a.Left <= b.Right && b.Left <= a.Right;
    internal static void Validate(WorksheetData data)
    {
        if (data.Validations.IsDefault || data.Validations.Length > 1024 || data.Validations.Any(v => v is null) ||
            data.Validations.Sum(v => v.Ranges.Length) > 4096 || data.Validations.Sum(v => v.EstimatedBytes) > 4 * 1024 * 1024)
            throw new ArgumentException("数据验证规则超过数量或文本容量限制。");
        var occupied = new List<CellRectangle>();
        foreach (var rule in data.Validations)
        foreach (var range in rule.Ranges)
        {
            if (occupied.Any(r => Overlaps(r, range))) throw new ArgumentException("数据验证区域不能重叠。");
            occupied.Add(range);
        }
    }
}
