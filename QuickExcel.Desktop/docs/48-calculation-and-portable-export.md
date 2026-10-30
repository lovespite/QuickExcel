# XC-0209 计算边界与可移植公式导出

2026-10-30，完成下述声明范围。沿用有界 AST、类型化计算器、静态依赖闭包和显式导出入口；没有扩大到整行、动态数组或 Excel 全函数。脚本仍由桌面的隔离 worker 执行。E2 下一项为 XC-0210。

## 67 个函数的参数契约

2026-10-31 增补：XC-0307 新增 SUBTOTAL，签名为 2–255 参数，当前共 68 个函数；五个真实 Excel 参数个数案例已追加到同一 fixture，原有 270 个案例保持。其引用、可见性和保留边界见 [56](56-outline-and-subtotals.md)。下表保留本项原 67 个函数的验收范围。

`StandardFunctions.AcceptsArguments` 是计算与导出共用的参数个数判定。Excel 16.0 build 20527、语言 2052 的 270 次真实公式赋值/拒绝记录固定在 `QuickExcel.Core.Test/Fixtures/excel-function-arity.json`。参数个数合法不代表该函数的所有数组、引用或文化模式已经实现；对应模式见 [44](44-conditional-functions.md)、[45](45-lookup-functions.md)、[46](46-date-and-math-functions.md)、[47](47-text-and-type-functions.md)。

| 函数 | 参数数 |
| --- | --- |
| SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, PRODUCT, AND, OR, CONCATENATE | 1–255 |
| SUMIF, AVERAGEIF, IF, MATCH, INDEX, WORKDAY, NETWORKDAYS, FIND, SEARCH | 2–3 |
| SUMIFS, AVERAGEIFS | 3–255，奇数，条件范围/条件成对 |
| COUNTIF, ROUND, IFERROR, EDATE, EOMONTH, ROUNDUP, ROUNDDOWN, MOD, TEXT, IFNA | 2 |
| COUNTIFS | 2–254，偶数，范围/条件成对 |
| COUNTBLANK, ABS, NOT, LEN, TRIM, UPPER, LOWER, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, INT, VALUE, ISBLANK, ISNUMBER, ISTEXT, ISERROR | 1 |
| LEFT, RIGHT, WEEKDAY | 1–2 |
| MID, DATE, TIME | 3 |
| VLOOKUP, HLOOKUP, SUBSTITUTE | 3–4 |
| XLOOKUP | 3–6 |
| XMATCH | 2–4 |
| REPLACE | 4 |
| TEXTJOIN | 3–254 |
| CONCAT | 1–254 |
| TODAY, NOW | 0 |

本轮将 CONCAT 上限从 253 修正为真实 Excel 验证的 254。Excel 接受 INDEX 的四参数引用形式，但该形式属于本产品未实现范围，导入走只读保留，不计入 2–3 参数标量支持。现代查找的省略参数按 [45](45-lookup-functions.md) 的 MissingNode 处理；普通 `IF(,...)`、`SUM(1,,2)` 仍不支持。

## 类型、错误和执行边界

| 场景 | 当前规则 |
| --- | --- |
| 普通引用 | 顶层公式引用空白格得到数值 0；在类型检测、拼接或比较等表达式内部仍保留 Blank 身份。单格矩形可作为标量，文本 ID 与错误外观文本保持 Text |
| SUM/AVERAGE/MIN/MAX/PRODUCT | 引用中的文本、布尔和空白不计入数值；直接及计算得到的标量按函数规则转换。没有数字时 AVERAGE 为 #DIV/0!，其余按既有规则得到 0 |
| COUNT/COUNTA | COUNT 计入数值及可转换的直接标量，忽略引用文本/布尔/错误；COUNTA 计入空文本和错误，忽略真实空白 |
| IF/INDEX 的引用返回 | 聚合和 AND/OR 保留被选引用的来源。例如文本 ID 引用经 IF 返回后仍不作为直接数字文本计数；只执行选中的 IF 分支一次 |
| IFERROR/IFNA 的空白返回 | 顶层结果为 0；ISBLANK 内仍可检测空白，CONCAT 内为空文本；在聚合/逻辑参数中按数值 0 处理。144 条专门的嵌套上下文 Excel 对照已固定 |
| 数字文本 | 算术及普通数字参数共用明确的有限数值语法，含已验证分组逗号、百分比、货币/符号和全角数字；不依赖进程文化。VALUE 的日期/时间扩展见 [47](47-text-and-type-functions.md)，日期函数仍按 [46](46-date-and-math-functions.md) 的规则 |
| 逻辑/比较 | 数字 0/非 0、布尔及明确 TRUE/FALSE 文本按逻辑规则处理；AND/OR 忽略引用文本/空白，但不是短路运算。比较不把数字外观文本自动改为数字 |
| 错误 | 普通函数/运算传播相应错误；COUNT、COUNTA、ISERROR 等有各自统计/检测规则。IFERROR 捕获错误，IFNA 仅捕获 #N/A。0^0 为 #NUM!、0 的负次幂为 #DIV/0!；非有限数字为 #NUM! |
| 懒分支 | IF、IFERROR、IFNA 只计算实际选择的值。依赖图仍收集全部静态分支；未选分支中的未知函数、禁用脚本或循环也可能导致保守受阻/冻结，不宣称动态依赖分析 |
| 文字/舍入 | 数字转文字使用 15 位有效数字；ROUND 修正十进制中点；LEFT/RIGHT 的大计数截到实际文本长度。MID 起始位置最多 2,147,483,648、长度最多 2,147,483,647，负小数按 Excel 边界拒绝。Unicode 版本与 32767 码元结果限制继续适用 |

标准函数不能被 UDF 同名覆盖。脚本结果类型和 v1/v2 桥接语义不变；真实 worker 回归验证嵌套 IF 聚合只调用选中的 NEXT 一次，不调用 FAIL。

## 引用和名称

整列按 Excel 1,048,576 行逻辑形状计算，实际内容仍稀疏存储，不能通过 INDEX 生成或编辑超出模型 100,000 行的地址。整行、3D、外部、联合/交叉引用、溢出与计算数组仍未实现。多格标量参数、IF 数组条件/结果、INDEX 的零轴数组或二维省略列号不会被误当成单值重算，导入保留原包与来源缓存。

INDEX 单格/行向量/列向量及矩形的标量结果已对照；多格维度的索引限于可判定的常量或 MATCH/XMATCH 结果。其他计算索引、单元格索引和无法确认数值的索引目前受保护，避免编辑后零索引变成未实现的数组结果。这是明确限制，不表示所有 Excel INDEX 模式均可计算。

名称绑定延续 [43](43-sheet-and-name-management.md)：单元格表达式局部优先，名称定义中未限定名称按工作簿作用域解析，显式表限定可选局部名称；单元格定义引用使用带表名的绝对地址。大小写不敏感，单引号/中文表名正确转义。递归名称、超过 16 层名称引用、16384 展开节点及 128 层 AST 深度受限。旧裸列拼写仅作为现存兼容行为，可移植导出会展开成标准整列引用。

## 普通保存与显式导出

普通 XLSX 保存保持公式，写入自动计算与 fullCalcOnLoad；来源不变仍可原包保留。重建 XLSX 时已知函数的非法参数个数（含名称定义）会在写文件前拒绝，目标文件保持不变；不把无法解释的 Excel 语法猜测为非法语法。XLSQ 仍可保留待修正公式。

“导出可移植公式”先显示保留、冻结、缺失和过期结果数量。受支持标准函数、可绑定跨表引用/名称、合法参数及有界标量语法可保留。脚本、未知/不支持语法、循环、非法参数以及名称展开后不能安全打印的公式作为冻结种子，连同跨表传递依赖转为类型化值。即使 IF 未选择脚本分支，也按静态闭包冻结。

名称全部展开并从输出中移除，避免隐藏 UDF。打印采用有界 StringBuilder，限制 8192 字符（含等号及 XLSX 现代函数前缀）、2048 节点/128 层并确认本产品能重新解析；超限不先分配无限展开字符串。冻结必须有结果；过期缓存必须经过现有明确确认。独立公式过期/缺失缓存会省略，交由 Excel 全重算。宿主 #CYCLE!/#SCRIPT! 映射为 Excel #VALUE!，#CALC! 继续使用富值结构。

输出按临时文件、重开校验、原子替换提交，允许已声明的行高单位舍入。没有脚本模块，保留六个独立标准公式的桌面样本通过 Excel 普通打开、全重算、保存重开。导出与取消均不修改源快照、脚本授权、撤销栈、脏状态或源文件字节。只读公式工作簿不能使用依赖范围未知的可移植导出，可选择既有显式缓存值快照导出。

## 验收记录

以下目录均在 `artifacts/compatibility`，不进入 Git。真实 Excel 为 16.0 build 20527 x64、语言 2052；仅操作任务自建文件和自有隐藏 COM 实例，宏/外链更新禁用，记录版本、来源、spec 和结果哈希。没有使用修复打开。

| 证据 | 结果 |
| --- | --- |
| `xc0209-excel-04` | 1352 条独立 Excel 标量/嵌套上下文公式，初始/编辑、重算及保存重开稳定 |
| `xc0209-probe-03`、`xc0209-verdict.json` | 10816 项模型初算/编辑/撤销重做/XLSX/XLSQ 重开检查，0 差异；8 类数组/整行保护原包不变；schema 0 问题；来源哈希不变 |
| `xc0209-output-initial-01`、`xc0209-output-edited-01` | 两份输出经 Excel 全重算及保存重开，1352 项逐字段一致，哈希绑定判定通过 |
| `xc0209-arity-02`、`xc0209-arity-03` | 270 个参数个数用例；固定 INDEX 引用形式的产品范围例外，CONCAT 上限修正后对照 0 差异 |
| `xc0209-context-01`、`xc0209-mid-01` | 144 个嵌套引用上下文、12 个 MID 精确整数上限对照；主矩阵已纳入上下文与关键边界 |
| `xc0209-tests-final` | 134 项计算、导出、存储及相关函数专项全部通过，含隔离 worker 懒分支、超长名称闭包、非法参数原子失败；未跑全量回归 |
| `xc0209-wpf-04` | 110 项实际公式栏/补全/撤销/导出处理器检查通过，取消与源文件/脏状态保护通过；导出前确认结果 Fresh |
| `xc0209-portable-oracle-01`、`xc0209-portable-excel-01` | 桌面 WPF 02 导出的 15 个值/公式检查格经独立 Excel 公式与全重算、保存重开对照通过；7 个冻结结果、6 个保留公式 |

固定输入/期望为 `calculation-boundary-spec.json`、`excel-calculation-boundary.json`，参数记录为 `excel-function-arity.json`。复现计算矩阵使用 [45](45-lookup-functions.md) 的 scalar-functions 命令，参数验收使用 `scripts/test-excel-formula-arity.ps1 -SpecPath QuickExcel.Core.Test/Fixtures/excel-function-arity.json -OutputDirectory artifacts/compatibility/<new>`；managed 命令一律指定 `--artifacts-path artifacts/build`。

失败证据保留：Excel 01 拒绝过大数字字面量后改为运算产生溢出；02 的 Formula2 产生 INDEX 数组标记，标量样本改为明确的 Legacy 赋值。probe 01/02 暴露数组保护缺口，03 修正。tests 01–04 固定空白/引用来源、数字文本、幂及 MID 边界；tests 06 暂时过宽的普通保存绑定检查拦截了有效但未实现的 TEXT 模式，已收窄为已知函数参数个数检查。WPF 01 的夹具用 v1 注册含点函数名被拒绝，改为真实 v2 描述符后通过。WPF 02/03 后补充视图刷新和 Fresh 断言，04 通过。

Release 解决方案构建通过，仅既有 NU1900 漏洞源连接警告。未发布新包，WPF 路由验收不等于物理输入或 OS IME，M4/M6/Beta 门禁继续保持。
