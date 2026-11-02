using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Printing;
using QuickExcel.Rendering.Skia;

internal static class CupsChecks
{
    internal static async Task<int> Run(string directory)
    {
        if(Directory.Exists(directory))return 2;
        // This test may only target the explicitly provisioned local test server.
        if(Environment.GetEnvironmentVariable("CUPS_SERVER")!="localhost:8631")throw new InvalidOperationException("CUPS test requires the isolated localhost:8631 server.");
        Directory.CreateDirectory(directory);var checks=new List<string>();Exception? failure=null;
        object? details=null;
        try
        {
            var printers=await CupsPrintService.GetDestinationsAsync();
            var printer=printers.Single(p=>p.Name=="QuickExcel-A6" && p.Instance is null);
            checks.Add("Real CUPS destination discovery");
            var sheet=new WorksheetSnapshot(SheetId.New(),"打印 中文",new Dictionary<CellPosition,CellSnapshot>{[new(0,0)]=new(CellContent.FromLiteral(CellValue.FromText("受控打印检查 中文 123")))});
            var book=new WorkbookSnapshot(Guid.NewGuid(),0,ExcelDateSystem.Windows1900,[sheet]);
            await using var session=await PrintSession.CreateAsync(book,sheet.Id,"controlled.xlsq",Path.Combine(directory,"printing"));
            var media=await CupsPrintService.ValidateMediaAsync(printer,session.Job.Layout);
            checks.Add("Printer media and hard margins match immutable PDF geometry");
            var rejected=false;try{await CupsPrintService.ValidateMediaAsync(printer,session.Job.Layout with {PaperWidth=123,PaperHeight=124});}catch(InvalidOperationException){rejected=true;}
            if(!rejected)throw new Exception("Unsupported paper was accepted.");checks.Add("Unsupported paper rejected before submitting any job");
            await session.ExportAsync(Path.Combine(directory,"expected.pdf"));
            await using var lease=session.RetainPdf();
            var receipt=await CupsPrintService.SubmitAsync(printer,media,session.PdfPath,"1",1);
            details=new{printer,media,receipt};checks.Add("Completed PDF handed to system spooler, receipt confirmed");
        }
        catch(Exception ex){failure=ex;}
        File.WriteAllText(Path.Combine(directory,"result.json"),JsonSerializer.Serialize(new{passed=failure is null,checks,details,error=failure?.ToString()},new JsonSerializerOptions{WriteIndented=true}));
        Console.WriteLine(failure?.ToString()??"PASS: controlled CUPS spool check");return failure is null?0:1;
    }
}
