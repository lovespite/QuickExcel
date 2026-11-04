using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation.Web;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptModuleTemplateTests
{
    [Fact]
    public void ModuleDescriptionIsBoundedSingleLineMetadata()
    {
        ScriptModuleSet Build(string description) => new(
            new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "export const value = 1;", Description: description) },
            ["main"], scriptApiVersion: 3);
        Assert.Equal("填充示例单元格", Build("填充示例单元格").Modules["main"].Description);
        Assert.Throws<ArgumentException>(() => Build(new string('x', ScriptModuleSet.MaximumDescriptionLength + 1)));
        Assert.Throws<ArgumentException>(() => Build("第一行\n第二行"));
    }

    [Fact]
    public void ValidModuleIdsProduceBoundedDistinctRegistrationNames()
    {
        string[] ids = ["123", "sample", "Sample", "sample-name", new('x', 128)];
        var modules = new ScriptModuleSet(ids.Select((id, index) => KeyValuePair.Create(id,
            new ScriptModule($"scripts/module{index}.js", ScriptModuleTemplates.Create(id, ScriptLanguage.JavaScript, true, true)))),
            ids, scriptApiVersion: 3, bindings: ids.SelectMany(id => new[] { new ScriptBinding(id, "formula-functions"), new(id, "actions") }));
        using var document = new WorkbookDocument(Book(Sheet("Data", (0, 0, F("=JS._123.DOUBLE(21)")))).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        document.EnableScriptsAndCalculate();
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
        Assert.Equal(ids.Length, document.ListActions().Count);
    }

    [Theory]
    [InlineData(ScriptLanguage.JavaScript, true, false)]
    [InlineData(ScriptLanguage.TypeScript, true, false)]
    [InlineData(ScriptLanguage.JavaScript, false, true)]
    [InlineData(ScriptLanguage.TypeScript, false, true)]
    [InlineData(ScriptLanguage.JavaScript, true, true)]
    [InlineData(ScriptLanguage.TypeScript, true, true)]
    public void StartersCalculateOrRunAtomicallyWithoutManualRepairs(ScriptLanguage language, bool functions, bool actions)
    {
        var source = ScriptModuleTemplates.Create("sample", language, functions, actions);
        var bindings = new List<ScriptBinding>();
        if (functions) bindings.Add(new("sample", "formula-functions"));
        if (actions) bindings.Add(new("sample", "actions"));
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        { ["sample"] = new($"scripts/sample.{(language == ScriptLanguage.TypeScript ? "ts" : "js")}", source, language) },
            ["sample"], scriptApiVersion: 3, bindings: bindings);
        var sheet = functions ? Sheet("Data", (0, 0, F("=JS.sample.DOUBLE(21)")), (0, 1, F("=JS.sample.DOUBLE(\"invalid\")"))) : Sheet("Data");
        using var document = new WorkbookDocument(Book(sheet).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        document.EnableScriptsAndCalculate();
        if (functions)
        {
            Assert.Equal("42", document.GetCellValue(0, 0, 0));
            Assert.Equal(CellValue.FromError(CellErrorCode.Script), document.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
            Assert.Contains(document.ScriptErrors, error => error.Contains("Expected a finite number."));
        }
        if (actions)
        {
            document.Session.UpdateNavigation(sheet.Id, new(1, 0));
            document.RunAction("sample");
            Assert.Equal("42", document.GetCellValue(0, 1, 0));
            Assert.True(document.Undo());
            Assert.Equal("", document.GetCellValue(0, 1, 0));
        }
        Assert.Same(modules.Modules, document.Snapshot.ScriptModules.Modules);
    }
}
