# XC-0205：条件统计公式

2026-10-30 完成声明范围。新增 SUMIF/SUMIFS、COUNTIF/COUNTIFS、AVERAGEIF/AVERAGEIFS、COUNTBLANK，接入内置函数登记、计算、补全/参数提示和既有公式保存。通过不可变快照及 DocumentSession 编辑，沿用撤销/重做、旧缓存身份和五秒计算预算，没有改变脚本权限或输入类型。

## 支持模式

| 函数 | 参数与范围规则 |
| --- | --- |
| SUMIF / AVERAGEIF | `range, criteria, [result_range]`；省略结果区域时使用条件区域；否则以结果区域左上角为起点，按条件区域的尺寸取值，原结果区域可更大或更小 |
| COUNTIF | `range, criteria`，计数满足条件的位置 |
| SUMIFS / AVERAGEIFS | 结果区域后跟 1–127 对条件区域/标量条件；每个区域必须同样大小和形状 |
| COUNTIFS | 1–127 对条件区域/标量条件，所有条件均满足才计数 |
| COUNTBLANK | 一个矩形区域；计入不存在的空格和公式返回的空文本 |

直接引用、跨表引用和绑定后的名称区域均支持，单个单元格也是合法区域。IFS 的参数配对、个数和形状错误返回 `#VALUE!`；无数字可平均时返回 `#DIV/0!`。结果区域的文本数字、布尔和空白不参与加总/平均；仅匹配位置的结果错误传播。隐藏/筛选状态不排除记录，SUBTOTAL 的不同口径属于 XC-0307。

SUMIF/AVERAGEIF 的投影区域同时用于依赖索引：原文本只写一个结果格时，隐式扩展覆盖的公式依赖和后续修改仍能触发重算；原文本写得更大时，未实际读取的尾部不产生错误循环。公式源码不因索引扩展而改写。

整列条件统计以 Excel 的 1,048,576 行为逻辑范围，未存空格用计数处理，不扩大产品的可存行列容量。只遍历条件/结果区域中已存位置的并集，最多 1,000,000 个稀疏偏移；超过预算给出受控错误，取消/超时不提交部分结果。结果起点偏移导致区域延伸到 Excel 网格外的部分视为空白。

## 条件类型与匹配

统一匹配器独立于普通公式比较，不把所有值转换为字符串：

- 支持 `= / <> / < / <= / > / >=`、大小写不敏感文本、`* / ?` 及 `~` 转义。包含通配符时处理转义；没有通配符的 `~~` 按字面匹配，已用真实 Excel 对照。
- 数字相等条件会匹配可解析的文本数字；数字不等与大小比较保持类型差异。例如数值 0 与文本 `"0"` 都匹配 `"=0"`，但文本 `"0"` 仍匹配 `"<>0"`。
- 空条件引用按数字 0 处理。条件 `""` 匹配真空格及空文本；`"="` 只匹配真空格，`"<>"` 包括空结果公式。只有比较运算符、没有操作数的大小条件不匹配任何值。
- 布尔和错误条件匹配对应类型，字符串 `"TRUE"` 或 `"#N/A"` 作为条件分别解释为布尔或错误，而不是同拼写文本。条件表达式自身产生错误时也可匹配该错误；结果区域的错误只在该位置被选中时传播。
- 条件文本超过 255 字符返回 `#VALUE!`。数字采用点号小数、合法三位逗号分组、科学计数和百分号。年在前的 `yyyy/M/d` / `yyyy-M-d` 公历日期按工作簿 1900/1904 序列解析；歧义月日顺序保持文本，虚构 1900-02-29 可用数值 60，不在本项文本日期解析范围。

文本排序使用 invariant 大小写无关比较；不承诺所有语言的拼音、宽窄字符或特殊排序与每种 Excel 区域设置相同。通配符采用无回溯匹配器，支持有界单元格文本。上述精确定义和固定样本优先于笼统“支持 COUNTIF”的说法。

标量条件可以是值、单格、名称或已有标量表达式。多格条件数组、INDEX 等计算得到的区域引用不在首期范围；名称展开后也检查模式。导入此类公式进入 XC-0106 的整簿只读/原包保留路径，缓存标为 Blocked；手工录入不支持模式产生明确公式错误。未把动态数组能力提前标记完成。

微软文档说明 [IFS 区域必须同形](https://support.microsoft.com/en-gb/excel/how-to-correct-a-value-error-in-the-sumif-sumifs-function) 和 [COUNTIF 条件/长度限制](https://support.microsoft.com/en-gb/excel/get-started/use-the-countif-function-in-microsoft-excel?nochrome=true)；类型、空文本、转义与错误细节以以下本机 Excel 独立实验固定，未只依赖文档概述。

## 证据

本机 Excel 16.0 build 20527、语言 2052、自动计算；使用自有隐藏 COM 实例，禁用宏及外链刷新，仅操作 artifacts 内自建文件。50 组条件 × 27 类输入的单格匹配，六类条件函数的全范围聚合，以及尺寸、整列、错误、转义等边界，共 1,706 条公式；另记录 Excel 直接拒绝录入的四个非引用区域公式。

| 证据（artifacts/compatibility） | 结果 |
| --- | --- |
| `xc0205-excel-oracle-final` | 1,706 条独立 Excel 公式在原始与编辑后均保存重开稳定；原始/编辑期望及版本、文化、计算模式、来源 SHA-256 完整记录 |
| `xc0205-probe-01` | 13,648 项：初算、编辑、一次撤销/重做、初始/编辑后的 XLSX/XLSQ 重开重算；0 差异，XLSX schema 0 问题，来源不变 |
| `xc0205-excel-initial-01`、`xc0205-excel-edited-01`、`xc0205-verdict-01.json` | 两份 QuickExcel 输出分别在 Excel 全重算并保存重开，公式、类型和值与独立预期精确一致；输出哈希绑定，0 差异 |
| `xc0205-related-final` | 126/126：条件函数、既有计算/存储/结构、公式辅助、受控公式、名称管理；含 127/128 对边界、依赖投影、增量复用、保存重开和数组条件保护 |
| `xc0205-wpf-01` | 49/49：既有录入辅助回归和全部七个新增函数的真实公式栏处理器、补全/提示、一次撤销/重做及普通 XLSX 保存重开；`conditional-entry.png` 已目视检查 |
| `xc0205-array-guard` | 早期真实 Excel 样本中的多格条件触发只读保护，来源缓存 Blocked，输出与原包逐字节相同，schema 0 问题 |

固定数值/错误期望保存在 `QuickExcel.Core.Test/Fixtures/excel-conditional-functions.json`，附原 oracle 哈希，测试不要求每次启动 Excel。模型浮点比较容差为相对 `1e-12`；最终实际 Excel 捕获的 JSON 精确比较，无容差替代。

复现：运行 `scripts/test-excel-conditional-functions.ps1 -OutputDirectory <新目录>`；兼容工具 `conditional-functions <oracle JSON> <新模型目录>`；脚本分别加 `-WorkbookPath <initial.xlsx / edited.xlsx> -OraclePath <oracle JSON> -State Initial/Edited`；工具 `verify-conditional-functions <oracle> <初始Excel JSON> <编辑Excel JSON> <模型报告> <新判定JSON>`。managed 命令仍使用 `--artifacts-path artifacts/build`；WPF 使用 `--formula-entry-smoke <新目录>`。

现代 Excel 源文件的未知部件保护不变；矩阵改写路径明确使用诊断副本，正常保存被拦截的原因保存在报告。自建桌面文件另行通过普通 XLSX 保存。早期 Excel oracle 01/02 因 Excel 拒绝非引用区域输入而失败，改为单独记录拒绝；Core 01 是统计属性拼写编译错误，Core 03 捕获了缺少操作数的大小条件错误匹配，已修复。旧证据不覆盖或删除。

Desktop Release 构建通过，仅既有 NU1900 漏洞源连接警告；未跑全量回归，未发布新包。界面为路由事件/实际处理器驱动，未宣称真实 OS IME 或物理按键验收；既有 M4/M6 门禁保留。
