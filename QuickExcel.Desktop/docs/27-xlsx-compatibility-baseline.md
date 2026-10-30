# XLSX 兼容基线与首轮证据

更新：2026-10-31。**XC-0001 已完成**：六份真实 Excel 业务样本、5384 属性断言和 E1 七项对照索引见 [37](37-excel-business-baseline.md)。XC-0002 的独立编辑预期与 48 路径对照已完成，见 [38](38-excel-edit-matrix.md)；XC-0003 亦已按 [39](39-xlsx-diagnostic-closure.md) 完成，E0 声明范围关闭。本文第 3–5 节保留初次切片的工具和历史证据，当时的“待真实 Excel”不覆盖后续 30–37 号记录；未发布新包。

## 1. 当前 R/D/C/E/W 矩阵

S = 支持该行限定子集；P = 部分支持；O = 仅原包保留；N = 不支持；拒绝 = 读取时拒绝；— = 不适用。此矩阵随验收更新，后文保留首轮历史证据。XC-0101–0107 的最新专项及真实 Excel 结果见 [30](30-excel-number-formats.md) 至 [36](36-xlsx-rich-text-links-and-notes.md)；W 已包含 XC-0103 接入的独立 XLSX 保存，值快照/可移植公式导出仍是不同操作。

| 特性与限定范围 | R | D | C | E | W | 证据/限制与下一项 |
| --- | --- | --- | --- | --- | --- | --- |
| 数值、文本编号、布尔、标准错误、空白/空文本 | S | S | S | S | S | `TypedStorageTests`、`roster`；不猜测数字外观文本的类型 |
| 1900/1904 日期原始序列和既有日期显示子集 | S | P | P | P | S | 两个 `date-ledger-*`；序列 60 原样保持；日期函数边界仍见 [10](10-m3-calculation.md) |
| 已有整数、两位小数、千分位、百分比 | S | S | — | P | S | `CellDisplayFormatter` 的明确格式子集；入口/效率扩展见 XC-0210 |
| 常见小数、货币/会计、数字分段/颜色、科学计数 | S | S | — | P | S | XC-0101 声明格式子集及 43 个 Excel 固定结果见 [30](30-excel-number-formats.md)，不支持任意 Excel 格式语法 |
| 基础字体、实色填充、边框/对齐、换行 | S | P | — | S | S | `TypedStyles` 与 [25](25-xlsx-rendering.md)；本轮修复字体 XML 子元素顺序 |
| 主题/索引/tint 颜色 | P | P | — | P | P | XC-0102 按验证后的原样式/主题关系保留，单格编辑追加明确覆盖，范围见 [31](31-xlsx-named-styles-and-theme.md) |
| Normal/命名样式、继承、apply 标志、主题字体 | P | P | — | P | P | 常用样式及主题关联、旋转/缩进/缩小填充已有 104 属性 Excel 对照，完整边界见 [31](31-xlsx-named-styles-and-theme.md) |
| 行列尺寸/隐藏/样式优先级 | P | P | — | P | P | 行列单格优先级与 customFormat 已有 [29 号专项证据](29-xlsx-row-style-inheritance.md)，新增 `row-column-styles` 第 10 个夹具；完整继承仍有限制；outline 已独立验收，见 [56](56-outline-and-subtotals.md) |
| 合并和基础浮动图片 | P | P | — | P | P | 既有 `TypedStorageTests`；复杂锚点/结构变化限制保持，XC-0402 |
| 普通/共享公式、跨表 A1、定义名称 | P | P | P | P | P | `cross-sheet`、`M3StorageTests`；支持函数与参数见 [10](10-m3-calculation.md)，缓存不等于重算 |
| 未知函数、`_xlfn` 前缀 | S | P | N | P | S | XC-0106：原公式/旧缓存保留，函数和依赖为 Blocked，独立公式可重算；结构变换拒绝，值/样式受限编辑可保存，见 [35](35-xlsx-preserved-formulas.md) |
| 数组公式 | S | P | N | N | O | XC-0106：整簿只读显示来源缓存，原 XLSX 包复制或 XLSQ 包装；不计算/编辑，见 [35](35-xlsx-preserved-formulas.md) |
| 结构化引用 | P | P | P | P | P | XC-0303 支持连续表列/当前行/特殊行、计算列、改名/插删/复制；20,891 项往返与 Excel/WPF 通过，边界见 [52](52-structured-references-and-table-editing.md) |
| 动态数组语义 | S | P | N | N | O | XC-0106 SEQUENCE 样本只读保留；计算/编辑待 XC-0501 |
| 文档属性白名单 | S | N | — | N | S | XC-0103 支持 core/app/custom 白名单；其他分组/结构仍报告，见 [32](32-xlsx-save-and-metadata.md) |
| 单窗口普通视图、冻结/缩放/活动表格、hidden/veryHidden | S | P | — | P | S | XC-0104 已持久化，视图命令可撤销、导航不 dirty；显示容量及拆分/多窗口限制见 [33](33-xlsx-workbook-views.md) |
| 页面/打印区域/重复标题/页眉页脚/分页 | S | N | — | N | S | XC-0105 的 R/W 子集；有分页时删除/移动行列拒绝，未提供实际打印或页面编辑，见 [34](34-xlsx-page-and-print-metadata.md) |
| 富文本 run、超链接、传统批注 | S | P | — | P | S | XC-0107 常见 run 字体/颜色、惰性链接说明及传统批注角标/悬停；显式降级与整格字体可撤销，结构/线程评论限制见 [36](36-xlsx-rich-text-links-and-notes.md) |
| 数据验证 | P | P | P | P | P | XC-0304 列表/范围/数值/日期/长度/基础公式、输入提示及整批校验；1,370 个真实 Excel 案例通过，范围见 [53](53-data-validation.md) |
| 条件格式基础 | P | P | P | P | P | XC-0305 比较/公式/重复值/文本、优先级与差异样式；真实 Excel 五阶段对照通过，结构与复制边界见 [54](54-conditional-formatting.md) |
| 高级条件格式、图表、透视表 | O | N | N | N | O | 色阶/数据条/图标集等保留/报告；图表/透视仍需 E4/E5 |
| 普通复制/TSV 粘贴、内部剪贴板/冻结 | — | P | — | P | P | 冻结 XC-0104、内部富剪贴板 XC-0201 已完成声明子集 |
| AutoFilter / sortState | P | P | — | P | P | XC-0301 多列列表/比较、排序设置和手动隐藏区分已持久化；37 个真实 Excel 案例及范围/来源限制见 [50](50-filter-and-sort-persistence.md) |
| Excel Table 基础 | P | P | — | P | P | XC-0302 名称/范围/列/表头/汇总/60 内置样式及独立筛选读显写、受限创建/底边调整；60,336 项往返和 Excel 复验见 [51](51-excel-table-basics.md)；结构化公式见 [52](52-structured-references-and-table-editing.md) |
| 常用清洗命令 | — | S | — | P | S | XC-0306 删除重复项、分隔/定宽分列、显式数字/日期类型；预览原子提交，双日期系统 Excel 对照和边界见 [55](55-data-cleaning.md) |
| 行列 outline 分组/折叠 | S | P | — | P | S | XC-0307 层级、隐藏/折叠、上下/左右方向及分组栏，复制/撤销/往返已验收；分组与活动筛选组合保留/报告，见 [56](56-outline-and-subtotals.md) |
| SUBTOTAL / 分类汇总生成 | P | P | P | P | P | 22 种模式的筛选/手动隐藏语义及 11 种预览生成；真实 Excel 14 表六阶段验证，引用/结构/混合状态边界见 [56](56-outline-and-subtotals.md) |
| 未知部件与关系 | O | N | — | N | O | 不变保留原 XLSX 字节；变更时可能阻止普通保存，不复制未知失效关系 |

矩阵的 S/P/O 不是互相替代的验收结论。对于一个含多种特性的文件，即使修改的单元格本身属于 S，其他特性的兼容问题仍可能阻止普通保存。未知函数的旧缓存已在后续报告中单独说明，但仍不能据此承诺安全重算或任意结构编辑。

## 2. 合成样本台账

所有样本来自 `tools/QuickExcel.CompatibilityProbe/FixtureCatalog.cs` 中独立构造的 SpreadsheetML，不经过 QuickExcel 写入器生成输入。原始 ZIP 与每次输出均落在 `artifacts/compatibility`，不纳入 Git。**这些文件不是 Excel 生成的样本**，未使用用户文件。

| ID | 场景/最小期望 | 对应项 |
| --- | --- | --- |
| `roster` | 中文名单、文本 `00123` 不变数值、布尔、空文本保持身份 | XC-0001、0107 |
| `quotation` | `1234.567`/`-42.5` 原始数值，货币分段代码、主题字体 scheme、applyFont=0 | XC-0101/0102 |
| `date-ledger-1900` | Windows1900，序列 60/0/45292.5 不发生时区或浮点类型转换 | XC-0101 |
| `date-ledger-1904` | 相同数值、Mac1904 标志保持；不把两种系统的同一数值当同一天 | XC-0101 |
| `cross-sheet` | `'数据'!A2*2`，输入 21、文件旧缓存 42；插行变 A3，改名引用跟随，删除引用变 #REF! | XC-0106/0209 |
| `print-template` | core 属性、冻结/125% 缩放/活动 B2、veryHidden、打印定义/页边距/方向/页眉页脚 | XC-0103/0104/0105 |
| `protected-print-template` | 同上并增加工作表保护；页面配置已支持，未实现的保护仍须报告并阻止有损普通保存 | XC-0404 |
| `unknown-function` | `_xlfn.XLOOKUP` 原始公式与旧缓存 7；本轮不执行重算 | XC-0106 |
| `array-formula` | `t=array` 与结果区域 B2:B3；XC-0106 后只读打开、缓存显示、原包保留，修改拒绝 | XC-0106 |
| `rich-links-notes` | 红色粗体 run、外部超链接、传统批注；常见子集普通格式保存通过，不访问外部地址 | XC-0107 |
| `phonetic-rich-text` | 已支持富文本上的未建模拼音，验证定位/容量及编辑后的保存保护 | XC-0107 |
| `row-column-styles` | 行/列/单格样式、显式行样式 0 与禁用行格式，验证继承和写回 | XC-0102 |

样本期望只是人工构造的诊断基准，不是 Excel 计算期望。每份报告明确 `Origin=synthetic-inline-spreadsheetml`、`ExcelVerification=pending-real-excel`；ExcelVersion/Build/Language/CalculationMode 为 null。未来补 Excel 生成且获准使用的脱敏样本时必须记录实际版本、语言、日期系统、计算模式、重开/修复提示及结果，不能用 SDK 的 Office2019 schema 目标充当实际 Excel 版本。

## 3. 可重跑工具

从仓库根执行：

```powershell
.\scripts\test-xlsx-compatibility.ps1
```

脚本只执行 `QuickExcel.CompatibilityProbe.Tests`，再生成独立基线目录，保存 TRX 和 `run-metadata.json`（基线提交、dirty 状态、SDK、校验目标）。不运行全量回归，不发布工具，不启用工作簿脚本。Open XML SDK 3.3.0 仅由开发工具引用；Desktop/Core/Infrastructure 没有新增 SDK 依赖。

每个样本独立从原文件打开，执行八条路径，归入计划的五类：不编辑、改 A2 值、切换 A2 粗体、插/删第 2 行或 B 列、首表改名。路径具体操作写入报告。普通保存生成 XLSQ，工具提取其中的 `workbook.xlsx` 供比较；尚未实现的产品 XLSX 保存入口不由此视为通过。

普通保存被兼容检查拒绝时，报告保留原错误，再在**新的诊断目录**生成名称带 `diagnostic-lossy-copy` 的受支持副本，用于揭示损失。不覆盖来源或旧证据；诊断副本不能作为普通保存成功。打开/命令/保存拒绝各有独立状态。无效命令或实质未改变内容也可能仍保留原包，不能把“路径有输出”理解为该特性可编辑。

只读比较任意两个获准检查的 XLSX：

```powershell
dotnet run --project .\tools\QuickExcel.CompatibilityProbe\QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path .\artifacts\build -- compare .\artifacts\before.xlsx .\artifacts\after.xlsx .\artifacts\comparison.json
```

比较结果包含双方哈希、独立结构检查、部件哈希、XML 元素/属性、公式、单元格值、样式及引用、关系、元数据差异。共享字符串按工作簿关系解析，数值比较使用 invariant 数值表示；不按固定部件文件名解析单元格。结构检查采用 `OpenXmlValidator(Office2019)`，另查内部目标、关系引用、样式/字符串索引。XML 禁止 DTD/外部实体，外部关系不连接；输入/解压总量上限各 64 MiB、2048 部件、XML 深度 64。

`compare` 无差异且无校验问题返回 0，有差异/校验问题返回 2，工具执行失败返回 1；已存在报告拒绝覆盖。`baseline` 的退出码 0 仅代表夹具有效且来源未改写，输出的既存兼容缺陷继续保留在报告，必须看 `summary.md` 和逐文件报告。

比较局限：XML 树位置、样式表重排及编码变化也会产生差异；不自动判断哪些差异属于预期编辑/等价表示，不能把差异数当作损失数。当前没有解析完整继承样式、展开共享公式来做等价证明，也没有像素比较/Excel 自动化。结构有效不代表公式正确或功能保真；本轮篡改公式用例专门验证“SDK 通过但语义差异仍能检出”。SDK/font 定义参考 [Microsoft Open XML Font 文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-3.0.1)。

## 4. 本轮结果与修复

最终证据：`artifacts/compatibility/e0-20261030-120601-ead41de4/`，专项 TRX 在同名前缀的 `-test-results/`。9 个输入均通过独立结构校验；72 条路径中 64 条生成输出，8 条数组公式路径在导入时拒绝；64 份输出均无结构校验错误，9 份来源哈希保持不变。17 条路径普通保存被阻止后生成了明确标注的有损诊断副本。这些数值不代表 64 次 Excel 兼容通过。

- `CompatibilityProbe.Tests`：21/21，通过删除关系/格式、篡改公式/打印元数据、缺失内部目标、类型保持、来源保护、拒绝覆盖、跨表变换，以及三种写入路径的字体 schema 回归。
- `QuickExcel.Core.Test` 仅筛选 `TypedStorageTests|M3StorageTests`：89/89。命令均使用 `--artifacts-path .\artifacts\build`；未跑全量套件。
- 本机 .NET SDK 10.0.401；NuGet 漏洞源不可达产生 NU1900 警告，包从本地缓存还原，编译和测试通过。
- Excel 实际版本/对照结果：未获得。本轮只查询到当前环境无 `Excel.Application` COM 注册，未据此断言机器上绝对没有安装 Excel，也未完成 Excel 重开/显示/重算验收。

首轮未修复证据保留在 `artifacts/compatibility/e0-20261030-120248-0ec78431/`：8 个输入有效，但 56 个输出中累计 100 条字体子元素顺序错误。`TypedStyles.Write` 原来按 name/sz/color/b/i 写入，读回器忽略顺序，因而内部校验漏检；现按 b/i/sz/color/name 写出。新测试分别覆盖 XLSQ 保存、值快照导出、可移植公式导出，验证模型中的字体值不变且独立结构校验通过。此修复只涉及 QuickExcel，未改动同级 ExcelIO。

## 5. 下一切片与关闭条件

XC-0001/0002 仍不勾选整体完成：需要扩充真实 Excel 来源样本与 E1 边界（主题字体完整映射、命名样式继承、分页/打印缩放等），增加编辑后期望差异分类和真实 Excel 复验；现有工具先用于驱动后续实现。

后续进展：XC-0003 首切片已将仅保留、内置不可计算、潜在损失及拒绝分开，补文件特性/当前公式定位，见 [28](28-xlsx-compatibility-diagnostics.md)。XC-0102 已修复行列样式优先级、customFormat 与空白格清除，10 个样本的最新基线和剩余边界见 [29](29-xlsx-row-style-inheritance.md)。继续按 XC-0101/0102/0103 的明确子集推进 P0 保真路径；禁止因减少提示而删除未知内容或绕过保存保护。本文第 4 节保留首轮历史数据，不覆盖 28/29 号记录的新增证据。
