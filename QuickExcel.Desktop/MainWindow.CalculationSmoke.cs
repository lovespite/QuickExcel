using System.IO;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCalculationSmokeAsync(string folder)
    {
        var path = Path.Combine(folder, "m3-ui.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(_services), path, "m3-ui.xlsq");
        ApplyValues([(0, 0, "7"), (0, 1, "=A1*2")]);
        _document.AddSheet("Summary");
        _document.SetCells(1, [(0, 0, "=Sheet1!B1+1"), (0, 1, "=SUM(Sheet1!A:A)")]);
        if (!await SaveAsync(false)) throw new Exception("M3 desktop save failed.");
        if (_document.GetCellValue(1, 0, 0) != "15") throw new Exception("M3 cross-sheet desktop calculation failed.");
        _document.SetCell(0, 0, 0, "9");
        if (!await WorkAsync("M3 增量验收…", _document.Calculate)) throw new Exception("M3 edit calculation failed.");
        if (_document.GetCellValue(1, 0, 0) != "19" || _document.GetCellValue(1, 0, 1) != "9")
            throw new Exception("M3 cross-sheet dependency update failed.");
        _document.Session.Execute(new RenameSheetCommand(_document.Snapshot.Sheets[0].Id, "数据表"));
        if (!await SaveAsync(false)) throw new Exception("M3 renamed save failed.");
        using var reopened = WorkbookDocument.Open(path, _services);
        reopened.Calculate();
        if (reopened.GetCellValue(1, 0, 0) != "19" || reopened.ScriptsEnabled)
            throw new Exception("M3 renamed workbook roundtrip failed.");
        var portable = Path.Combine(folder, "m3-portable.xlsx");
        reopened.ExportPortableXlsx(portable);
        var exported = new TypedWorkbookStorage().Open(portable);
        if (exported.Sheets[1].GetCell(0, 0)!.Content.Kind != CellContentKind.Formula)
            throw new Exception("M3 portable export discarded independent formula.");
        File.WriteAllText(Path.Combine(folder, "calculation-smoke-result.txt"),
            "PASS: desktop composition uses typed cross-sheet calculation; input edit updates transitive closure; renamed references save/reopen; portable formulas retained; reopened scripts disabled.");
    }
}
