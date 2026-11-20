using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Storage;

if (args.Length != 1) throw new ArgumentException("Pass the repository samples directory.");
var directory = Path.GetFullPath(args[0]);
Directory.CreateDirectory(directory);
const string csv = "商品,数量,单价\r\n\"苹果,红\",2,3.5\r\n香蕉,3,2";
Seed[] seeds = [
    new("lodash", "Lodash · 订单去重与分组", "Orders: deduplicate and group", "按订单号去重，按类别汇总金额并排序。", "Deduplicate orders, group totals by category and sort.",
        [["订单号","类别","金额"],["A001","办公",120],["A002","办公",80],["A003","食品",35],["A003","食品",35],["A004","食品",65]],
        [["类别","去重订单数","金额合计"],["办公",2,200],["食品",2,100]], "A2:C6", "去重前 5 条记录，去重后 4 笔订单。"),
    new("decimal.js", "decimal.js · 精确金额计算", "Exact decimal amounts", "保留文本单价精度，计算小计和金额合计。", "Calculate exact subtotals and totals from text prices.",
        [["商品","单价（文本）","数量"],["笔芯","0.10",3],["笔记本","19.99",3],["试算","1.005",2]],
        [["商品","精确单价","数量","小计（字符串）"],["笔芯","0.10",3,"0.30"],["笔记本","19.99",3,"59.97"],["试算","1.005",2,"2.01"],["合计","","","62.28"]], "A2:C4", "单价和计算结果使用字符串，避免先转 Number 丢失精度。"),
    new("dayjs", "Day.js · 日期与时区", "Dates and time zones", "严格校验日期，转换 UTC 并计算七天后日期。", "Validate dates, convert to UTC and add seven days.",
        [["时间（文本）","时区"],["2024-02-29 09:30","Asia/Shanghai"],["2024-07-01 12:00","America/New_York"],["2024-02-30 08:00","Asia/Shanghai"]],
        [["输入时间","时区","UTC 时间","七天后日期"],["2024-02-29 09:30","Asia/Shanghai","2024-02-29 01:30 UTC","2024-03-07"],["2024-07-01 12:00","America/New_York","2024-07-01 16:00 UTC","2024-07-08"],["2024-02-30 08:00","Asia/Shanghai","无效日期","—"]], "A2:B4", "UTC、时区和严格解析插件已经预装；日期采用固定输入，不依赖当前日期。"),
    new("papaparse", "Papa Parse · CSV 解析与导出", "Parse and export CSV", "解析含引号、逗号和换行的 CSV，再生成 CSV 文本。", "Parse quoted CSV with commas and line breaks, then export it.",
        [["CSV 文本"],[csv]],
        [["商品","数量","单价","小计"],["苹果,红",2,3.5,7],["香蕉",3,2,6],["","","",""],["导出 CSV",csv,"",""]], "A2", "编辑 A2 的多行 CSV；本例只处理内存文本，不下载文件。"),
    new("simple-statistics", "simple-statistics · 统计与回归", "Statistics and regression", "计算均值、中位数、标准差、相关性与线性预测。", "Compute descriptive statistics, correlation and linear predictions.",
        [["X","Y"],[1,3],[2,5],[3,7],[4,9],[5,11]],
        [["指标","数值"],["样本数",5],["Y 均值",7],["Y 中位数",7],["Y 总体标准差",2.828427],["回归斜率",2],["回归截距",1],["相关系数",1],["X=6 的预测值",13]], "A2:B6", "输入满足 Y=2X+1，标准差保留六位小数。"),
    new("zod", "Zod · 表格数据校验", "Validate table data", "逐行检查姓名、邮箱和年龄，显示可读错误信息。", "Validate names, emails and ages with readable row errors.",
        [["姓名","邮箱","年龄"],["Alice","alice@example.com",28],["Bob","bad-email",-1],["","carol@example.com",""]],
        [["原始行号","姓名","校验结果","详情"],[2,"Alice","通过","28 岁 · alice@example.com"],[3,"Bob","未通过","email: 邮箱格式无效；age: 年龄至少为 1"],[4,"","未通过","name: 姓名不能为空；age: 年龄至少为 1"]], "A2:C4", "safeParse 返回校验结果，不会因为一行无效而中断其他行。"),
    new("fuse.js", "Fuse.js · 商品模糊搜索", "Fuzzy product search", "使用带拼写错误的搜索词查找商品。", "Find products even when the search term has a typo.",
        [["商品","类别","价格","","搜索词"],["Apple iPhone","手机",5999,"","iphne"],["Apple iPad","平板",2999],["Google Pixel","手机",4999],["Samsung Galaxy","手机",4599],["Apple MacBook","电脑",8999]],
        [["商品","类别","价格"],["Apple iPhone","手机",5999]], "A2:C6、E2", "把 E2 改为 ipad 或 galaxy，再运行以观察搜索结果。"),
    new("arquero", "Arquero · 表格分组聚合", "Group and aggregate tables", "将单元格数据转为表，分组计数、求和并排序。", "Create a table, group rows, aggregate amounts and sort.",
        [["类别","地区","金额"],["办公","华东",120],["办公","华北",80],["食品","华东",35],["食品","华北",65],["食品","华东",25]],
        [["类别","记录数","销售合计"],["办公",2,200],["食品",3,125]], "A2:C6", "本例使用内存表操作，不需要网络权限。"),
    new("mathjs", "math.js · 数学运算", "Expressions, units and matrices", "演示表达式、单位换算、符号求导、矩阵与大数。", "Evaluate expressions, convert units, differentiate and multiply matrices.",
        [["参数","值","","矩阵 A","","","矩阵 B",""],["表达式","sqrt(16)+2","",1,2,"",5,6],["长度","2 inch","",3,4,"",7,8],["函数","x^2"]],
        [["运算","结果"],["表达式",6],["英寸转厘米",5.08],["对 x 求导","2 * x"],["矩阵乘积 (1,1)",19],["矩阵乘积 (1,2)",22],["矩阵乘积 (2,1)",43],["矩阵乘积 (2,2)",50],["大数 0.1 + 0.2","0.3"]], "B2:B4、D2:E3、G2:H3", "修改表达式、单位或矩阵；单元格只能存放标量，所以矩阵结果逐项写入。")
];
var storage = new TypedWorkbookStorage();
var catalog = new List<object>();
foreach (var (seed, index) in seeds.Select((value,index)=>(value,index)))
{
    var instructions = Table(seed.Id, "使用说明", [
        [seed.Title], [seed.Description], ["使用方法"],
        ["1. 查看“输入数据”和“预期结果”，也可以修改输入。"],
        ["2. 打开“脚本 → 脚本编辑器”，选择 Actions 下的 example 查看代码。"],
        ["3. 点击“应用并重算”（F5），按正常流程同意本次脚本执行。"],
        ["4. 在“自动执行”中选择 example，点击“执行”，查看“运行结果”。"],
        ["5. 可撤销本次运行；需要保留修改时另存自己的工作簿。"],
        ["输入范围：" + seed.InputRange], [seed.Note],
        ["预期结果对应出厂输入，不会跟随修改自动变化。"],
        ["所有示例均可离线运行，打开示例不会自动授权或执行脚本。"]], instructions:true);
    var input = Table(seed.Id, "输入数据", seed.Input);
    var output = Table(seed.Id, "运行结果", [["请应用脚本并运行 example；结果将显示在这里。"]]);
    var expected = Table(seed.Id, "预期结果", seed.Expected);
    var modules = new ScriptModuleSet(new Dictionary<string,ScriptModule> {
        ["main"] = new("scripts/main.ts", "// 初始化入口；应用脚本只加载 example 操作。\nexport function main(workbook: QuickExcel.Workbook): void {}\n",ScriptLanguage.TypeScript),
        ["common"] = new("scripts/common.ts",Source("common"),ScriptLanguage.TypeScript),
        ["example"] = new("scripts/example.ts",Source(seed.Id),ScriptLanguage.TypeScript,Description:seed.Description)
    }, ["main","example"],scriptApiVersion:4,bindings:[new("main","initialization"),new("example","actions")]);
    var workbook = new WorkbookSnapshot(Id(seed.Id),0,ExcelDateSystem.Windows1900,[instructions,input,output,expected],styles:[
        new(), new() { FontSize=11, WrapText=true, Vertical=1 },
        new() { FontSize=12, Bold=true, FillColor="FFE4F1E9", FontColor="FF185C37", WrapText=true, Vertical=1 }
    ]).WithScriptModules(modules);
    var file = $"{index+1:00}-{seed.Id}.xlsq";
    storage.Save(Path.Combine(directory,file),workbook);
    catalog.Add(new { id=seed.Id, file, title=seed.Title, titleEn=seed.TitleEn, description=seed.Description, descriptionEn=seed.DescriptionEn,
        sha256=Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(directory,file)))) });
    Console.WriteLine(file);
}
File.WriteAllText(Path.Combine(directory,"catalog.json"),JsonSerializer.Serialize(new { schemaVersion=1, samples=catalog },new JsonSerializerOptions { WriteIndented=true, Encoder=System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping })+"\n",new UTF8Encoding(false));

string Source(string id) => File.ReadAllText(Path.Combine(directory,"scripts",id+".ts")).Replace("\r\n","\n");
static Guid Id(string text) => new(SHA256.HashData(Encoding.UTF8.GetBytes("QuickExcel.Sample/"+text)).AsSpan(0,16));
static WorksheetSnapshot Table(string id,string name,object[][] values,bool instructions=false)
{
    var cells = new Dictionary<CellPosition,CellSnapshot>();
    for(var row=0;row<values.Length;row++) for(var column=0;column<values[row].Length;column++)
    {
        var value=values[row][column];
        if(value is string { Length:0 })continue;
        var cell=value is string text ? CellValue.FromText(text) : CellValue.FromNumber(Convert.ToDouble(value,System.Globalization.CultureInfo.InvariantCulture));
        cells.Add(new(row,column),new(CellContent.FromLiteral(cell,row==0?2:1)));
    }
    return new(new(Id(id+"/"+name)),name,cells,
        rows:Enumerable.Range(0,40).Select(row=>KeyValuePair.Create(row,new RowMetadata(row==0?32:instructions?32:id=="papaparse"&&(row==1||row==4)?80:42))),
        columns:Enumerable.Range(0,instructions?1:8).Select(column=>KeyValuePair.Create(column,new ColumnMetadata(instructions?105:id=="papaparse"?38:column==3&&id=="zod"?65:id=="dayjs"?29:24))),
        defaultRowHeight:24,tabColor:name=="输入数据"?"FF5588BB":name=="运行结果"?"FF237A57":null);
}
record Seed(string Id,string Title,string TitleEn,string Description,string DescriptionEn,object[][] Input,object[][] Expected,string InputRange,string Note);
