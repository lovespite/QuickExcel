using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class M4PolicyTests
{
    [Theory]
    [InlineData("Date.now()")]
    [InlineData("Math.random()")]
    [InlineData("Function('return 7')()")]
    [InlineData("eval('7')")]
    public void ConservativeWorkerStillExposesSourcesThatPreventAPurityClaim(string expression)
    {
        using var runtime = new ScriptWorkerRuntimeFactory().Create($"host.functions.register('F',()=>{expression});");
        Assert.Equal("conservative", Assert.Single(runtime.FunctionDescriptors).Profile);
        Assert.IsType<double>(runtime.Invoke("F", []));
    }

    [Fact]
    public void UnsupportedStrictDescriptorIsRejectedByRealWorker()
    {
        Assert.Equal("Load", Assert.Throws<ScriptRuntimeFailureException>(() => new ScriptWorkerRuntimeFactory().Create(
            "host.functions.register({name:'JS.F',apiVersion:2,profile:'strict'},()=>1);")).Code);
    }

    [Fact]
    public void HiddenMutableModuleStateFollowsWorkerLifetimeAndIsNeverDeduplicatedOrCached()
    {
        var sheet = Sheet("S", (0, 0, F("=NEXT(1)")), (1, 0, F("=NEXT(1)")), (2, 0, F("=A1+A2")));
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", "import {next} from 'state'; host.functions.register('NEXT',next);"),
            ["state"] = new("scripts/state.js", "let hidden=0; export const next=x=>x+(++hidden);")
        }, ["main"], scriptApiVersion: 3);
        using var session = new DocumentSession(Book(sheet).WithScriptModules(modules));
        session.SetScriptsEnabled(true);
        var calculator = new WorkbookCalculator(new ScriptWorkerRuntimeFactory());
        for (var i = 0; i < 2; i++)
        {
            var result = calculator.Calculate(session.BeginCalculation());
            var offset = (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()) ? i * 2 : 0;
            Assert.Equal(CellValue.FromNumber(2 + offset), C(result, sheet).Value);
            Assert.Equal(CellValue.FromNumber(3 + offset), C(result, sheet, 1).Value);
            Assert.Equal(CellValue.FromNumber(5 + offset * 2), C(result, sheet, 2).Value);
            Assert.Equal(0, result.Statistics!.ReusedCount);
            Assert.True(session.TryCommitCalculation(result));
        }
    }
}
