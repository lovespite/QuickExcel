using System.Security.Cryptography;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

internal static class PivotProbe
{
    internal static int Run(string source,string directory)
    {
        source=Path.GetFullPath(source); directory=Path.GetFullPath(directory);
        if (Directory.Exists(directory)) throw new InvalidOperationException("Use a new evidence directory.");
        Directory.CreateDirectory(directory); var bytes=File.ReadAllBytes(source); var hash=SHA256.HashData(bytes); var storage=new TypedWorkbookStorage();
        using var session=new DocumentSession(storage.Open(source));
        var pivots=session.Snapshot.Sheets.SelectMany(s=>s.Data.Pivots).ToArray();
        var pivotIssues=session.Snapshot.FileState!.CompatibilityDiagnostics.Where(d=>d.Feature.StartsWith("worksheet/pivot",StringComparison.OrdinalIgnoreCase) ||
            d.Feature.StartsWith("relationship/pivot",StringComparison.OrdinalIgnoreCase) || d.Feature.StartsWith("part/",StringComparison.OrdinalIgnoreCase) && d.Feature.Contains("pivot",StringComparison.OrdinalIgnoreCase)).ToArray();
        ProbeRunner.WriteNew(Path.Combine(directory,"import.json"),JsonSerializer.SerializeToUtf8Bytes(new { Count=pivots.Length, PivotIssues=pivotIssues,
            OtherIssues=session.Snapshot.FileState.CompatibilityDiagnostics.Except(pivotIssues).ToArray(), Pivots=pivots.Select(p=>new{p.Name,p.Source,p.Range,p.Stale,Records=p.Cache.Records.Length}) },new JsonSerializerOptions { WriteIndented=true }));
        if (pivots.Length!=2 || pivotIssues.Length!=0) throw new InvalidOperationException("Independent Excel pivot import did not fully model both basic reports; inspect import.json.");
        var sheet=session.Snapshot.Sheets.Single(s=>s.Name=="Data");
        session.Execute(new SetCellsCommand([new(new(sheet.Id,1,2),CellContent.FromLiteral(CellValue.FromNumber(12)))]));
        var stale=Path.Combine(directory,"stale.xlsx"); storage.Save(stale,session.Snapshot,discardUnsupported:true);
        foreach (var report in session.Snapshot.Sheets)
        foreach (var pivot in report.Data.Pivots) session.Execute(new RefreshPivotCommand(report.Id,pivot.Name));
        var refreshed=Path.Combine(directory,"refreshed.xlsx"); storage.Save(refreshed,session.Snapshot,discardUnsupported:true);
        var opened=storage.Open(refreshed);
        if (opened.Sheets.SelectMany(s=>s.Data.Pivots).Any(p=>p.Stale) || !SHA256.HashData(File.ReadAllBytes(source)).SequenceEqual(hash))
            throw new InvalidOperationException("Refresh or source preservation failed.");
        Console.WriteLine("PASS: two independently generated Excel pivots imported, marked stale, refreshed and rewritten; source unchanged.");
        return 0;
    }
}
