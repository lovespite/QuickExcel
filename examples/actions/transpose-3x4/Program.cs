using System.Globalization;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;

// 从仓库根目录运行；生成文件和所有构建产物均保存在 artifacts。
if (!File.Exists("QuickExcel.slnx"))
    throw new InvalidOperationException("请从 QuickExcel 仓库根目录运行此示例。");

var path = Path.GetFullPath("artifacts/examples/transpose-3x4.xlsq");
Directory.CreateDirectory(Path.GetDirectoryName(path)!);
var source = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "main.js"));
var modules = new ScriptModuleSet(
    new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source) },
    ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")]);
var services = new WorkbookServices(new TypedWorkbookStorage(),
    new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor());

using (var seed = WorkbookDocument.CreateBlank(services))
{
    seed.SetCells(0, Enumerable.Range(0, 12).Select(i =>
        (i / 4, i % 4, (i + 1).ToString(CultureInfo.InvariantCulture))));
    seed.Session.Execute(new UpdateScriptModulesCommand(modules));
    seed.Save(path);
}

// 打开生成的包，通过真实 worker 验证动作、一次撤销与重做。
// 验证期间不保存，交付的文件始终保留运行前的 3×4 数据。
using var document = WorkbookDocument.Open(path, services);
if (document.ScriptsEnabled) throw new InvalidOperationException("打开文件不应自动启用脚本。");
document.EnableScriptsAndCalculate();
var action = document.ListActions().Single();
var revision = document.Snapshot.Revision;
if (!document.RunAction(action.Name).Changed || document.Snapshot.Revision != revision + 1)
    throw new InvalidOperationException("转置未作为一次事务提交。");
Verify([
    [1, 5, 9, null],
    [2, 6, 10, null],
    [3, 7, 11, null],
    [4, 8, 12, null]
]);
document.DisableScripts();
if (!document.Undo() || document.CanUndo) throw new InvalidOperationException("一次撤销失败。");
Verify([
    [1, 2, 3, 4],
    [5, 6, 7, 8],
    [9, 10, 11, 12],
    [null, null, null, null]
]);
if (!document.Redo()) throw new InvalidOperationException("重做失败。");
Verify([
    [1, 5, 9, null],
    [2, 6, 10, null],
    [3, 7, 11, null],
    [4, 8, 12, null]
]);
Console.WriteLine("PASS: package roundtrip, worker transpose, cleared source tail, one undo/redo.");
Console.WriteLine(path);

void Verify(int?[][] expected)
{
    for (var row = 0; row < expected.Length; row++)
    for (var column = 0; column < expected[row].Length; column++)
    {
        var value = document.Snapshot.Sheets[0].GetCell(row, column)?.Value ?? CellValue.Blank;
        var wanted = expected[row][column] is { } number ? CellValue.FromNumber(number) : CellValue.Blank;
        if (value != wanted) throw new InvalidOperationException($"Unexpected value at ({row}, {column}).");
    }
}
