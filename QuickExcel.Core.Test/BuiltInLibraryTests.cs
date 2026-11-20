using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Core;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class BuiltInLibraryTests(Xunit.Abstractions.ITestOutputHelper output)
{
    public static IEnumerable<object[]> Libraries => new[] { "lodash", "decimal.js", "dayjs", "papaparse", "simple-statistics", "zod", "fuse.js", "arquero", "mathjs" }.Select(name => new object[] { name });

    [Theory]
    [MemberData(nameof(Libraries))]
    public void BuiltInFunctionsAndDocumentedUnsupportedOperations(string name)
    {
        var cases = JsonNode.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures/ScriptLibraryCases.json")))![name]!.AsObject();
        var expected = JsonNode.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures/ScriptLibraryExpected.json")))![name]!.AsObject();
        var expressions = cases.Select(c => "(()=>{try{return {id:" + JsonSerializer.Serialize(c.Key) + ",value:(()=>{" + c.Value!.GetValue<string>() + "})()}}catch(e){return {id:" + JsonSerializer.Serialize(c.Key) + ",error:String(e)}}})()");
        using var runtime = Runtime($"import * as lib from '{name}'; export function CHECK() {{return JSON.stringify([{string.Join(',', expressions)}]);}}");
        var results = JsonNode.Parse((string)runtime.Invoke("JS.main.CHECK", [])!)!.AsArray();
        Assert.Equal(cases.Count, results.Count);
        foreach (var result in results)
        {
            var id = result!["id"]!.GetValue<string>();
            var match = expected[id]!;
            if (match["error"] is { } error) Assert.Contains(error.GetValue<string>(), result["error"]?.GetValue<string>() ?? "");
            else Assert.True(JsonNode.DeepEquals(match["value"], result["value"]) && result["error"] is null, $"{name}/{id}: {result}");
        }
    }

    public static IEnumerable<object[]> TypedImports => new[] {
        new[] { "lodash", "import _, {sumBy} from 'lodash'; const n:number=sumBy([{v:20},{v:22}],'v');", "n" },
        new[] { "decimal.js", "import Decimal from 'decimal.js'; const n:string=new Decimal('0.1').plus('0.2').toString();", "n==='0.3'?42:0" },
        new[] { "dayjs", "import dayjs from 'dayjs'; const n:number=dayjs.duration({minutes:42}).asMinutes(); const t:string=dayjs.tz('2024-01-01','Asia/Shanghai').utc().format();", "n" },
        new[] { "papaparse", "import Papa, {parse} from 'papaparse'; const rows=parse<{n:number}>('n\\n42',{header:true,dynamicTyping:true}).data; const text:string=Papa.unparse(rows);", "rows[0].n" },
        new[] { "simple-statistics", "import {mean} from 'simple-statistics'; const n:number=mean([40,44]);", "n" },
        new[] { "zod", "import {z} from 'zod'; const n:number=z.number().parse(42); const u:string=z.url().parse('https://example.com'); const f:File=z.file().parse(new File(['x'],'a.txt'));", "n" },
        new[] { "fuse.js", "import Fuse from 'fuse.js'; const matches=new Fuse([{name:'apple',n:42}],{keys:['name']}).search('aple'); const n:number=matches[0].item.n;", "n" },
        new[] { "arquero", "import {from,op} from 'arquero'; const n:number=from([{n:20},{n:22}]).rollup({n:op.sum('n')}).get('n',0) as number;", "n" },
        new[] { "mathjs", "import {add,unit} from 'mathjs'; const n:number=add(20,22); const cm:number=unit('2 inch').toNumber('cm');", "n" }
    }.Select(row => row.Cast<object>().ToArray());

    [Theory]
    [MemberData(nameof(TypedImports))]
    public void IsolatedCompilerChecksRealUpstreamTypesAndEmittedImports(string name, string setup, string result)
    {
        Assert.Contains(name,setup);
        using var runtime = Runtime(setup + " export function CHECK():number {return " + result + ";}", ScriptLanguage.TypeScript);
        Assert.Equal(42d, runtime.Invoke("JS.main.CHECK", []));
    }

    [Fact]
    public void InvalidLibraryTypesAndUndeclaredPathsAreRejected()
    {
        foreach (var source in new[] {
            "import Decimal from 'decimal.js'; const n:number=new Decimal('1');",
            "import {sumBy} from 'lodash'; const n:number=sumBy([{v:1}], 'v'); const bad:string=n;",
            "import {missing} from 'zod'; export const value=missing;",
            "import 'https://example.com/a.js';",
            "import 'dayjs/plugin/not-declared';"
        }) Assert.Throws<ScriptCompilationException>(() => Runtime(source, ScriptLanguage.TypeScript));
    }

    [Theory]
    [InlineData("lodash\0extra")]
    [InlineData("./lodash")]
    [InlineData("Lodash")]
    [InlineData("node:fs")]
    public void RuntimeImportsOnlyExactCatalogIds(string id)
    {
        var error=Assert.Throws<ScriptRuntimeFailureException>(()=>Runtime("import * as lib from "+JsonSerializer.Serialize(id)+"; export function CHECK(){return 1;}"));
        Assert.Contains("declared package module ID",error.Message);
    }

    [Fact]
    public void AllLibrariesShareOneWorkerWithoutBecomingUdfEntrypoints()
    {
        const string source = """
            import _ from 'lodash'; import Decimal from 'decimal.js'; import dayjs from 'dayjs';
            import Papa from 'papaparse'; import * as ss from 'simple-statistics'; import {z} from 'zod';
            import Fuse from 'fuse.js'; import * as aq from 'arquero'; import * as math from 'mathjs';
            export function CHECK() {
              const rows=Papa.parse<{name:string,n:number}>('name,n\napple,20\nbanana,22',{header:true,dynamicTyping:true}).data;
              const n=aq.from(rows).rollup({total:aq.op.sum('n')}).get('total',0);
              if(new Fuse(rows,{keys:['name']}).search('aple')[0].item.n!==20)throw Error('search');
              if(ss.mean(rows.map(r=>r.n))!==21 || _.sumBy(rows,'n')!==42)throw Error('statistics');
              if(dayjs.utc('2024-01-01').year()!==2024 || Number(math.add(20,22))!==42)throw Error('date/math');
              return z.number().parse(new Decimal(n as number).toNumber());
            }
            """;
        using var runtime = Runtime(source,ScriptLanguage.TypeScript);
        Assert.Equal(42d, runtime.Invoke("JS.main.CHECK", []));
        Assert.Equal(42d, runtime.Invoke("JS.main.CHECK", []));
    }

    [Fact]
    public void BuiltinsDoNotConsumeUserModuleCountOrSourceQuotaAndLocalIdsKeepPrecedence()
    {
        var modules = Enumerable.Range(0, 63).ToDictionary(i => "helper" + i, i => new ScriptModule("scripts/helper" + i + ".js", "export const n=1;"));
        const string source = "import {sum} from 'lodash'; export function CHECK(){return sum([20,22]);}";
        var size = modules.Values.Sum(m => System.Text.Encoding.UTF8.GetByteCount(m.Source));
        modules["main"] = new("scripts/main.js", source + "/*" + new string(' ', 1024 * 1024 - size - source.Length - 4) + "*/");
        using var runtime = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(new ScriptModuleSet(modules,["main"],scriptApiVersion:4),new(Guid.NewGuid(),0,0,0),TimeSpan.FromSeconds(5),default));
        Assert.Equal(42d,runtime.Invoke("JS.main.CHECK",[]));
        using var local = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(new ScriptModuleSet(new Dictionary<string,ScriptModule> {
            ["main"]=new("scripts/main.js","import {n} from 'lodash'; export function CHECK(){return n;}"),
            ["lodash"]=new("scripts/local.js","export const n=7;") },["main"],scriptApiVersion:4),new(Guid.NewGuid(),0,0,0),TimeSpan.FromSeconds(5),default));
        Assert.Equal(7d,local.Invoke("JS.main.CHECK",[]));
        using var typedLocal = new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(new ScriptModuleSet(new Dictionary<string,ScriptModule> {
            ["main"]=new("scripts/main.ts","import {n} from 'dayjs'; import {mean} from 'simple-statistics'; export function CHECK():number{return mean([n,n]);}",ScriptLanguage.TypeScript),
            ["dayjs"]=new("scripts/local.ts","export const n=7;",ScriptLanguage.TypeScript) },["main"],scriptApiVersion:4),new(Guid.NewGuid(),0,0,0),TimeSpan.FromSeconds(5),default));
        Assert.Equal(7d,typedLocal.Invoke("JS.main.CHECK",[]));
    }

    [Fact]
    public void LibraryActionsCommitAndRollbackThroughTheExistingTransaction()
    {
        var modules = new ScriptModuleSet(new Dictionary<string,ScriptModule> {
            ["main"]=new("scripts/main.ts", """
                import {sumBy} from 'lodash'; import {z} from 'zod';
                export function main(workbook:QuickExcel.Workbook) {
                  const sheet=workbook.getActiveWorksheet();
                  sheet.getRange('A1').setValue(sumBy([{n:20},{n:22}],'n'));
                  z.number().positive().parse(sheet.getRange('B1').getValue());
                }
                """,ScriptLanguage.TypeScript) },["main"],scriptApiVersion:4,bindings:[new("main","actions")]);
        using var doc = new WorkbookDocument(Book(Sheet("Data", (0,1,N(1)))).WithScriptModules(modules),
            new(new TypedWorkbookStorage(),new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),new ScriptWorkerActionExecutor()));
        doc.EnableScriptsAndCalculate(); doc.RunAction("main"); Assert.Equal("42",doc.GetCellValue(0,0,0));
        Assert.True(doc.Undo()); Assert.Equal("",doc.GetCellValue(0,0,0));
        doc.SetCell(0,0,1,"-1");
        Assert.ThrowsAny<Exception>(()=>doc.RunAction("main")); Assert.Equal("",doc.GetCellValue(0,0,0));
    }

    [Fact]
    public async Task ArqueroLoadsThroughAuthorizedHostFetchAndCollectsUtf8Streams()
    {
        await using var server = new ScriptFetchTestServer(_ => Task.FromResult(new ScriptFetchTestServer.Response(200,"name,n\n你好,42")));
        var modules = new ScriptModuleSet(new Dictionary<string,ScriptModule> {
            ["init"]=new("scripts/init.js","host.permissions.network.request('127.0.0.1');"),
            ["main"]=new("scripts/main.js",$$$"""
                import * as aq from 'arquero';
                export async function main(w) {
                  const table=await aq.loadCSV('{{{server.Url}}}data.csv');
                  if(table.get('name',0)!=='你好')throw Error('network UTF-8');
                  const data=new TextEncoder().encode('name,n\n你好,42');
                  let offset=0,released=false;
                  const stream={getReader(){return {
                    async read(){return offset<data.length?{done:false,value:data.slice(offset,++offset)}:{done:true};},
                    async cancel(){}, releaseLock(){released=true;}
                  };}};
                  const split=await aq.fromCSVStream(stream);
                  if(!released||split.get('name',0)!=='你好')throw Error('split UTF-8');
                  const json=await aq.fromJSONStream(new Response('[{"n":42}]').body);
                  const arrow=await aq.fromArrowStream(new Response(table.toArrowIPC()).body);
                  const fixed=await aq.fromFixedStream(new Response('ab42').body,{widths:[2,2],names:['name','n']});
                  if(json.get('n',0)!==42||arrow.get('n',0)!==42||fixed.get('n',0)!==42)throw Error('stream roundtrip');
                  let sent=false;
                  const text=await aq.fromCSVStream({getReader(){return {async read(){if(sent)return {done:true};sent=true;return {done:false,value:'name,n\n文本,42'};},async cancel(){},releaseLock(){}};}});
                  if(text.get('name',0)!=='文本')throw Error('text stream');
                  let cancelled=false,unlocked=false;
                  try {
                    await aq.fromCSVStream({getReader(){return {async read(){return {done:false,value:new Uint8Array(10*1024*1024+1)};},async cancel(){cancelled=true;},releaseLock(){unlocked=true;}};}});
                    throw Error('oversized stream accepted');
                  } catch(e) {if(!(e instanceof RangeError)||!cancelled||!unlocked)throw e;}
                  try {await aq.loadCSV('https://unauthorized.invalid/a.csv.gz');throw Error('compression accepted');}
                  catch(e) {if(!String(e).includes('codec'))throw e;}
                  let blocked=false; try{await aq.loadCSV('https://unauthorized.invalid/data.csv')}catch(e){blocked=true;}
                  if(!blocked)throw Error('network consent bypass');
                  w.getActiveCell().setValue(split.get('n',0));
                }
                """) },["init","main"],scriptApiVersion:4,bindings:[new("init","initialization"),new("main","actions")]);
        using var doc=new WorkbookDocument(Book(Sheet("Data")).WithScriptModules(modules),
            new(new TypedWorkbookStorage(),new WorkbookCalculator(new ScriptWorkerRuntimeFactory()),new ScriptWorkerActionExecutor()));
        doc.Session.Network.Authorize=(hosts,_)=>hosts.SequenceEqual(["127.0.0.1"]);
        doc.EnableScriptsAndCalculate();doc.RunAction("main");Assert.Equal("42",doc.GetCellValue(0,0,0));
        Assert.True(doc.Undo());Assert.Equal("",doc.GetCellValue(0,0,0));
    }

    private IScriptRuntime Runtime(string source, ScriptLanguage language = ScriptLanguage.JavaScript)
    {
        try { return new ScriptWorkerRuntimeFactory().Create(new ScriptRuntimeRequest(new ScriptModuleSet(new Dictionary<string,ScriptModule> {
            ["main"]=new("scripts/main."+(language==ScriptLanguage.TypeScript?"ts":"js"),source,language)
        },["main"],scriptApiVersion:4),new(Guid.NewGuid(),0,0,0),TimeSpan.FromSeconds(5),default));
        } catch (ScriptCompilationException e) { output.WriteLine(string.Join("\n", e.Diagnostics.Select(d=>d.Message))); throw; }
    }
}
