# XC-0208：文本和类型判断公式

2026-10-30 完成声明范围。TEXT/VALUE、SUBSTITUTE/REPLACE、FIND/SEARCH、TEXTJOIN/CONCAT、ISBLANK/ISNUMBER/ISTEXT/ISERROR/IFNA 共 13 个函数接入类型化计算、依赖追踪、补全/参数提示、撤销及 XLSX/XLSQ。输入类型契约不变；文本转换仅在函数内部发生。

## 函数范围

| 函数 | 已固定行为 |
| --- | --- |
| TEXT | 复用 XC-0101 数字/日期格式解析与渲染，支持数字、百分比、科学计数、分节、条件、字面量和日期/时间；格式参数为常量或直接/名称绑定的单格文字/数值输入 |
| VALUE | 明确采用点号小数、逗号数字分组、科学计数、单个前/后百分号、`$`/`¥` 前缀及括号负数；支持全角数字归一化、年在前的日期和 24 小时时间文本。布尔、空文本、歧义日期及未声明拼写返回 `#VALUE!` |
| SUBSTITUTE | 区分大小写；全部替换非重叠，指定第 n 次时按 Excel 计入重叠命中；空 old_text 原样返回；instance_num 向下取整且须大于 0 |
| REPLACE | 以 1 为起点；起点/长度向下取整，分别限 1…32767 / 0…32767；超出实际文本后追加，长度截到末尾 |
| FIND | 区分大小写、按字面查找；空查找文本允许指向末尾后一位 |
| SEARCH | 不区分大小写，支持 `*`、`?` 和 `~` 转义；空来源以及超出来源末尾的起点返回 `#VALUE!` |
| CONCAT / TEXTJOIN | 按参数顺序、区域内按行连接，支持直接/名称区域、跨表和整列；TEXTJOIN 明确处理 ignore_empty、真空白、空字符串及分隔符；数值文本采用 15 位有效数字 |
| ISBLANK / ISNUMBER / ISTEXT / ISERROR | 以原值类型判断；`""` 是文本而非真空白，数字文本不是 Number，错误拼写文本不是 Error |
| IFNA | 只捕获 `#N/A`，保留其他错误；未选择分支不执行，选择真空白引用返回数值 0 |

TEXT 格式语言固定为本机验收的简体中文 Excel 语言，不随操作系统 CurrentCulture 变化。常规格式使用 `G/通用格式`；英文 `General` 和 `[Red]` 等颜色名在 TEXT 参数中按实测返回 `#VALUE!`，单元格 XLSX 数字格式仍使用英文标准格式码。支持已验证的中文颜色和 `[颜色1]`…`[颜色56]`，颜色仅影响格式选择，不成为返回文本的一部分；本机不接受的 `[青色]` 等拼写保持错误。日期月/星期名默认英文，可沿用已支持的明确区域格式标记。1900 序列 0/60 与前两月星期沿用 Excel 的历史规则；负日期在 1900 系统返回错误，1904 系统按实测的负号/绝对序列格式化。

VALUE 的日期支持 `yyyy-M-d` / `yyyy/M/d` 和可选时间；独立时间为 `H:m[:s[.fraction]]`，小时不超过 9999，分秒小于 60，不接受 AM/PM 或地区月/日猜测。数字分组按本机 Excel 实测接受每个逗号后至少三位；分组前须非零起始，`1,2345` 可转换，`1,23` 和 `0,123` 不可。括号负数不能再有正负号，货币与百分比不能组合。不把此有限语法宣称为所有地区的 Excel VALUE。

返回文本限 32767 个 UTF-16 码元。连接函数先传播输入错误，再判断连接长度；在本机 Excel build 20527，超长 CONCAT/TEXTJOIN 为 `#CALC!`，SUBSTITUTE/REPLACE 超长仍为 `#VALUE!`。稀疏整列只遍历存储格并按间隔计算空白分隔符，不分配百万个空单元格；查找/替换/连接保留取消与既有计算时间预算。

TEXT 的分数/DBNum 等未实现格式、计算得到的格式参数、多格标量参数和 IF 等计算得到的多格连接参数在绑定时明确拒绝。导入使用整簿只读、Blocked 旧缓存和逐字节原包保存；未开放动态数组。TEXT 格式引用若本身是公式，也属于该保护范围。

## 文本兼容版本与文件表示

Excel 的[兼容版本说明](https://support.microsoft.com/en-gb/excel/compatibility-versions)明确：旧工作簿使用版本 1；版本 2 改变 LEN、MID、FIND、SEARCH、REPLACE 的 Unicode 代理对计数，组合标记、变体选择符和肤色修饰符仍独立计数。QuickExcel 新建空簿设为 2，无扩展的历史文件设为 1；工作簿模型、所有相关命令、增量缓存和 XLSX/XLSQ 都携带版本。LEFT/RIGHT 在本机两种版本中均按完整 Unicode 字符截取。

版本 1 可以截出单独的 UTF-16 代理码元。写入文本/缓存时用 XLSX `_xHHHH_` 转义，读取后恢复原码元，保留完整代理对和原有字面转义语义。支持 `workbookCompatibilityVersion` 的 setVersion 1/2；未来版本的公式只读保留，非法字段拒绝，尚未建模的警告阈值/附加属性继续报告损失。受支持副本使用当前版本作为警告阈值。XLSQ 必需能力 `text-compatibility-v2` 防止旧读取器改变计算语义。

`#CALC!` 采用微软 [MS-XLSX Calc Error Type](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/a4cdc4c3-45b4-45c4-be49-d244f598a538) 的 `_error` / `errorType=13`：单元格写 `#VALUE!` 回退值及 vm 引用，配套 sheetMetadata、rich value、structure 三个部件和关系；读取时恢复独立 Calculation 错误身份。不能把 `#CALC!` 直接写到普通 `<v>` 或作为 Excel 公式错误字面量。错误身份通过 XLSQ `extended-cell-errors` 能力声明保护；模型的纯值导出也使用富值表示。原文件的额外错误说明、其他富值类型及元数据操作策略仍受诊断保护，不静默舍弃。

IFNA/CONCAT/TEXTJOIN 在 XLSX 使用已实现的 `_xlfn.` 别名，模型使用规范名称；字符串字面量、工作表引号内容和未知函数前缀不重写。

## 验收证据

Excel 16.0 build 20527 x64，语言 2052。仅使用任务自建隐藏 COM 实例和 artifacts 文件，关闭宏和外链更新；记录进程归属、版本、来源与 spec 哈希。支持矩阵每个日期系统 2,329 条公式，另有八个保护文件。

| artifacts/compatibility 下证据 | 结果 |
| --- | --- |
| `xc0208-excel-1900-04`、`xc0208-excel-1904-04` | 合计 4,658 条独立 Excel 公式，初始/编辑、全重算、保存重开稳定 |
| `xc0208-probe-1900-03`、`xc0208-probe-1904-03` | 合计 37,264 项初算、编辑、撤销/重做和 XLSX/XLSQ 重开检查，0 差异；16 个保护样本只读、保留旧缓存与原包；schema 0 问题 |
| `xc0208-output-1900-initial-02` / `xc0208-output-1900-edited-01`、1904 同名目录；`xc0208-verdict-1900.json` / `1904.json` | 四份模型输出均被 Excel 普通打开，全重算、另存重开与独立预期精确一致，哈希绑定判定通过 |
| `xc0208-version-1-oracle-02`、`xc0208-version-2-oracle-01` | 各九条真实 Excel Unicode/超长连接对照，版本差异固定；版本 1 源 XML 的 `_xDE00_`/`_xD83D_` 与 Core 码元级往返测试对应 |
| `xc0208-tests-02`、`xc0208-tests-03`、`xc0208-tests-final` | 182 项显示/视图/文本相关回归、213 项计算/存储/命令相关回归、最终 34 项文本/版本/富值异常文件专项全部通过；集合有交集，不累加为独立测试总数 |
| `xc0208-wpf-02` | 99/99：补全、提示、实际公式栏提交、一次撤销/重做、普通 XLSX 重开；`text-type-entry.png` 已目视核对 |
| `xc0208-ui-oracle-01`、`xc0208-ui-excel-01` | 桌面普通保存的 13 个新函数经 Excel 全重算及保存重开，逐字段一致 |

固定 spec/期望位于 `QuickExcel.Core.Test/Fixtures/text-type-1900-spec.json` / `text-type-1904-spec.json` 及 `excel-text-type-*.json`，期望附 oracle SHA-256。复现使用 [45](45-lookup-functions.md) 的 scalar-functions 工具与 `scripts/test-excel-scalar-functions.ps1`；显式 TextCompatibilityVersion 可生成版本 1/2 的验收工作簿。managed 命令均带 `--artifacts-path artifacts/build`。

失败证据保留：初版发现 IFNA 空白返回差异；version-1-oracle-01 的 PowerShell XAttribute 字符串转换导致重复扩展，已改用 Value；大矩阵 02 验证了 Excel 不接受 `#CALC!` 公式字面量，03/专项 01 固定 VALUE 分组/符号/时间上界、TEXT 颜色、输入错误优先级。WPF 01 因新增 SUBSTITUTE 改变 `su` 候选排序而失败，测试改为明确选择 sum。模型 probe 02 的直接 `#CALC!` 错误写法被 Excel 正常 Open 拒绝（output initial 01），最终 probe 03 使用标准富值结构后通过；未使用修复打开掩盖格式问题。

现代 Excel 源文件的其他未知部件保护保持：矩阵改写使用明确诊断副本，并记录普通保存拦截；桌面自建文件使用普通保存。Release 构建通过，仅既有 NU1900 漏洞源连接警告。未跑每项全量回归，未发布新包；路由事件烟测不等于物理输入/OS IME，既有 M4/M6 门禁不因此关闭。下一项 XC-0209。
