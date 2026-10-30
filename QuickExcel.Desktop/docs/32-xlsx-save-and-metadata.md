# XLSX 保存与文档属性保留

2026-10-30，XC-0103 完成，提交以本文件首次出现的提交定位。范围为公式保真的 XLSX 保存路径、低耦合属性白名单、受影响部件计划与既有保存保护；不是任意 Excel 文件的无损编辑承诺。

## 保存契约

- `Save` / `SaveChecked` 接受 XLSX 与 XLSQ。桌面打开 XLSX 后保留文件路径和打开指纹，Ctrl+S 原位保存；另存提供“Excel 工作簿（保留公式）”。保存包含公式原文和现有缓存，不调用结果导出或可移植公式冻结逻辑。重写设置 `fullCalcOnLoad=1`，重开缓存仍为待重算，不把旧缓存称为新结果。
- 两种格式均先写同目录临时文件并重开验证，再检查目标指纹和原子替换。XLSX 校验值、公式、样式、名称、工作表与已建模配置；文档属性另按保留计划验证字节、结构和关系。校验/冲突失败不替换目标，不清除 dirty。XLSX 不承诺保存 XLSQ 的文档/表稳定身份和脚本修订。
- 未编辑且模型与原文件相同，可整包逐字节保留。内容改变时，未知部件仍阻止普通保存；明确选择受支持副本只允许另存，原文件不覆盖。保存后当前会话仍保留原来源及诊断，后续编辑可能再次提示；重新打开副本后使用副本自身的来源清单。
- 新建空白簿和导入 XLSX 的脚本为空。演示簿仍带示例脚本；已有 XLSQ 的脚本原样保留。只允许规范的空 API v1 单模块保存为 XLSX；非空内容（包括注释/空格）、模块配置、动作或签名都要求 XLSQ 主文件，禁用执行也不等于删除脚本。界面自动转入 XLSQ 另存，标题说明主文件与 Excel 副本；底层同样拒绝，`discardUnsupported` 不能绕过。
- 两个现有导出操作继续分别表示“所有公式转为缓存值”和“保留支持公式、显式冻结其余公式及依赖”。导出 XLSX 不改变主文件路径，不标记脚本主文件已保存。

## 元数据范围与依赖

`XlsxMetadata` 根据根关系、内容类型、命名空间、元素/属性、标量类型和无出站关系的条件验证属性。不会枚举并复制全部未知 ZIP 条目。属性只读保留（R/W），本项没有属性编辑界面。

| 部件 | 支持白名单 | 编辑时处理 |
| --- | --- | --- |
| Core properties | 标题、主题、作者、描述、标识、语言；关键字、修订、分类、状态、版本、上次作者；合法 XML 日期时间 | 与单元格/表结构无依赖，字节保留 |
| Extended properties | Application、AppVersion、Company、Manager、Template、HyperlinkBase；DocSecurity 及 ScaleCrop/LinksUpToDate/SharedDoc/HyperlinksChanged 标量 | 静态属性保留；标志不表示已经实现对应协作或安全功能 |
| HeadingPairs / TitlesOfParts | 单一 `Worksheets` 或 `工作表` 分组，计数、类型和标题必须与源工作表完全一致 | 改名、增删、重排后更新计数和标题；其他分组/语言保持诊断，不能只按碰巧相同的字符串推断 |
| Custom properties | 无链接属性，唯一 name/pid，标准 fmtid；字符串、布尔、int/i4/i8、有限 r8、合法 date/filetime | 字节保留，文本编号不转数字；链接、blob、复杂类型和未知字段不复制 |
| 已建模配置 | 日期系统、网格线、默认行高、行列尺寸/隐藏、标签色、普通名称；此前样式/主题子集 | 由当前模型及 XC-0102 的关联规则重建/保留 |

任一属性部件含未知结构、依赖或不合法标量时，整个部件报告为仅保留；改动后的普通保存阻止，显式兼容副本会移除已报告的该部件。没有悄悄丢掉单个未知属性后仍宣称该部件完整。计算配置只接受已知引擎编号/重算标志，迭代、手动计算等其他设置进入诊断。

`InspectSavePlan` 输出整包保留、模型部件重建、样式/主题验证保留、属性保留、标题索引更新及未知内容阻止/显式副本的判定。写入共用属性计划，写后另验证计划。结构不明的内容没有通用“保留即可安全编辑”的豁免。

语义参考：[Microsoft extended properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.properties)、[HeadingPairs 分组与计数](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.headingpairs)、[Excel 自定义文档属性](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.customdocumentproperties)。

## 验证与复现

本机 Excel 16.0 build 20527，64 位，语言 2052，自动计算（-4105）；脚本只处理新建隔离夹具和新输出，不访问用户工作簿。Excel 以 `xlNormalLoad` 打开，没有请求修复；本轮不使用不存在的 `RepairMode` 属性作为证据。

| 检查 | 结果 / `artifacts/compatibility` 下证据 |
| --- | --- |
| 保存、属性、脚本、冲突、既有 XLSQ / facade 回归 | 125/125，`xc0103-regressions-final`，14:13:32 TRX |
| 新增写后校验失败的目标保护 | 保存专项最终 18/18，`xc0103-save-tests-final`，14:14:53 TRX |
| 独立 Open XML 比较工具回归 | 34/34，`xc0103-probe-tests` |
| WPF 实际打开/保存/另存处理器 | 50/50（本项新增 14），`xc0103-wpf-03/xlsx-rendering-smoke-result.txt`；取消/冲突/明确副本/脚本转存与 dirty 验证，对话框选择注入，非物理输入 |
| Excel 来源文件 | `xc0103-excel-07/excel-source.xlsx` 与 `excel-metadata.json` |
| 导入、改值、改表名、直接保存 XLSX、重开后再次原位保存 | `xc0103-metadata-roundtrip-02/metadata-report.json`：来源哈希不变、core/custom 字节不变、app 静态属性不变、输出独立 schema 校验 0 问题 |
| Excel 读取 QuickExcel 输出、全量重算、Excel 保存重开 | `xc0103-excel-reopen-02/excel-metadata.json`：7 项内置/3 项自定义属性与来源一致，`00123` 保持文本；A1=9，B1 仍为 `=SUM(A1,3)`，结果/显示 12，改名保留；再次保存重开相同 |

自定义浮点费率在 Excel 自身保存过程中从 0.15 变为 0.15000000000000002；原始数值完整留在报告，比较容差明确为 1e-12。早期脚本 COM 调用、属性索引与扩展名断言失败证据保留在先前编号目录，不作为通过结果。

复现入口：`scripts/test-excel-metadata.ps1 -OutputDirectory <新 artifacts 目录>` 创建 Excel 来源；`dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- metadata <来源.xlsx> <新目录>` 生成直接 XLSX 保存证据；再运行脚本并指定 `-WorkbookPath <输出.xlsx> -ReferencePath <来源 excel-metadata.json>` 做真实 Excel 属性对照与重算重开。

普通 Excel 来源仍带视图、打印、修订扩展、calcChain 等本项未建模内容，所以真实来源的首次改写采用明确的兼容副本，损失清单完整保留；该副本重开后可普通原位保存。冻结/缩放/活动表由 XC-0104、打印设置由 XC-0105 接续；复杂公式的受控读取/结构编辑由 XC-0106 接续。不能把这次属性与保存验证说成普通模板所有内容已无损。

没有升级 XLSQ 格式版本：属性存在于内嵌的标准 XLSX，旧读取器按原未知部件机制保护；脚本格式/API/权限未改变。Desktop Release 构建通过；仅有既有 NU1900 漏洞源不可达警告。未运行完整回归、性能测试或发布包装，未更新旧发布包，E0/E1 与 E3 均未整体关闭。
