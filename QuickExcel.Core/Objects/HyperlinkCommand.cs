using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Edits inert link metadata without changing the cell text, style or other attachments.</summary>
public sealed class SetHyperlinkCommand(SheetId sheetId, CellPosition position, CellHyperlink? link) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookLimits.ValidatePosition(position.Row, position.Column);
        var sheet = snapshot.GetSheet(sheetId); var old = sheet.Annotations;
        if (old.Links.GetValueOrDefault(position) == link) return new(snapshot, [], false);
        if (link is not null)
        {
            if (!string.IsNullOrEmpty(link.Target))
            {
                if (!Uri.TryCreate(link.Target, UriKind.Absolute, out var uri) || uri.Scheme is not ("https" or "http" or "mailto"))
                    throw new ArgumentException(MessageText.Get("链接目标须为 http、https、mailto 地址，或使用工作簿内位置。"));
            }
            else if (string.IsNullOrWhiteSpace(link.Location) ||
                FormulaBinding.Bind(FormulaSyntax.Parse("=" + link.Location), snapshot, sheetId, position) is not ReferenceNode)
                throw new ArgumentException(MessageText.Get("内部链接须引用现有工作表地址或区域名称。"));
        }
        var links = link is null ? old.Links.Remove(position) : old.Links.SetItem(position, link);
        var annotations = new WorksheetAnnotations(links, old.Notes, old.NoteDrawing);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithAnnotations(annotations) : s)), [], false, CalculationUnchanged: true);
    }
}
