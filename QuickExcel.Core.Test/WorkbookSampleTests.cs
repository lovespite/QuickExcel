using System.Text.Json.Nodes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Calculation;
using QuickExcel.Runtime;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class WorkbookSampleTests(Xunit.Abstractions.ITestOutputHelper output)
{
    private static string Samples => Path.Combine(AppContext.BaseDirectory,"samples");

    [Theory]
    [InlineData("lodash")]
    [InlineData("decimal.js")]
    [InlineData("dayjs")]
    [InlineData("papaparse")]
    [InlineData("simple-statistics")]
    [InlineData("zod")]
    [InlineData("fuse.js")]
    [InlineData("arquero")]
    [InlineData("mathjs")]
    public void ShippedWorkbookRunsRealTypedExampleMatchesExpectedAndCanUndo(string id)
    {
        var catalog=new WorkbookSampleCatalog(Samples);
        Assert.Equal(9,catalog.Samples.Count);
        var entry=Assert.Single(catalog.Samples,s=>s.Id==id);
        var path=Path.Combine(Samples,entry.File);
        var original=File.ReadAllBytes(path);
        var snapshot=catalog.Read(id);
        Assert.Equal(4,snapshot.ScriptModules.ScriptApiVersion);
        Assert.Null(snapshot.ScriptModules.Signature);
        Assert.Equal(File.ReadAllText(Path.Combine(Samples,"scripts",id+".ts")).Replace("\r\n","\n"),snapshot.ScriptModules.Modules["example"].Source);
        Assert.Equal(File.ReadAllText(Path.Combine(Samples,"scripts/common.ts")).Replace("\r\n","\n"),snapshot.ScriptModules.Modules["common"].Source);
        using var document=new WorkbookDocument(snapshot,new(new TypedWorkbookStorage(),new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),new ScriptWorkerActionExecutor()));
        Assert.False(document.ScriptsEnabled);
        var before=document.GetCellInput(2,0,0);
        try
        {
            document.EnableScriptsAndCalculate();
            Assert.Equal(before,document.GetCellInput(2,0,0));
            Assert.Equal("example",Assert.Single(document.ListActions()).Name);
            document.RunAction("example");
        }
        catch(ScriptCompilationException error) { output.WriteLine(string.Join("\n",error.Diagnostics.Select(d=>d.Message))); throw; }
        AssertResults(document.Snapshot);
        Assert.True(document.Undo());
        Assert.Equal(before,document.GetCellInput(2,0,0));
        Assert.True(document.Redo()); AssertResults(document.Snapshot);
        document.RunAction("example"); AssertResults(document.Snapshot);
        Assert.Equal(original,File.ReadAllBytes(path));
    }

    private static void AssertResults(WorkbookSnapshot snapshot)
    {
        var expected=snapshot.Sheets.Single(s=>s.Name=="预期结果");
        var actual=snapshot.Sheets.Single(s=>s.Name=="运行结果");
        foreach(var position in expected.Cells.Keys.Union(actual.Cells.Keys))
        {
            var wanted=expected.Cells.GetValueOrDefault(position)?.Value??CellValue.Blank;
            var value=actual.Cells.GetValueOrDefault(position)?.Value??CellValue.Blank;
            if(wanted.Kind==CellValueKind.Number)Assert.Equal(wanted.NumberValue,value.NumberValue,8);
            else if(wanted.Kind==CellValueKind.Blank)Assert.True(value.Kind==CellValueKind.Blank||value is { Kind:CellValueKind.Text, TextValue:"" });
            else Assert.Equal(wanted,value);
        }
    }

    [Fact]
    public void ShorterCsvRerunClearsPreviousRowsBeyondTheInitialDisplayArea()
    {
        var snapshot=new WorkbookSampleCatalog(Samples).Read("papaparse");
        using var document=new WorkbookDocument(snapshot,new(new TypedWorkbookStorage(),new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),new ScriptWorkerActionExecutor()));
        var csv="商品,数量,单价\n"+string.Join('\n',Enumerable.Range(1,65).Select(i=>$"Item{i},2,3"));
        document.SetCell(1,1,0,csv);document.EnableScriptsAndCalculate();document.RunAction("example");
        Assert.Equal("Item65",document.GetCellValue(2,65,0));
        document.SetCell(1,1,0,"商品,数量,单价\nOnly,1,2");document.RunAction("example");
        Assert.Equal("Only",document.GetCellValue(2,1,0));
        Assert.Equal("",document.GetCellValue(2,65,0));
        Assert.Equal("",document.GetCellValue(2,67,1));
        Assert.True(document.Undo());Assert.Equal("Item65",document.GetCellValue(2,65,0));
    }

    [Fact]
    public void CatalogRejectsEscapingPathsAndChangedWorkbookBeforeOpening()
    {
        var directory=Path.Combine(AppContext.BaseDirectory,"test-data","samples-"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var content=JsonNode.Parse(File.ReadAllText(Path.Combine(Samples,"catalog.json")))!;
            var entry=content["samples"]![0]!;
            var file=entry["file"]!.GetValue<string>();
            entry["file"]="../outside.xlsq";
            File.WriteAllText(Path.Combine(directory,"catalog.json"),content.ToJsonString());
            Assert.Throws<InvalidDataException>(()=>new WorkbookSampleCatalog(directory));
            entry["file"]=file;
            File.WriteAllText(Path.Combine(directory,"catalog.json"),content.ToJsonString());
            File.WriteAllText(Path.Combine(directory,file),"changed package");
            var catalog=new WorkbookSampleCatalog(directory);
            Assert.Throws<InvalidDataException>(()=>catalog.Read(entry["id"]!.GetValue<string>()));
            Assert.Throws<ArgumentException>(()=>catalog.Read("unknown"));
        }
        finally { Directory.Delete(directory,true); }
    }
}
