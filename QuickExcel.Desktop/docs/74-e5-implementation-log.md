# E5 高级分析与扩展兼容实施记录

2026-11-04：XC-0501–0504 的实现与专项验收已按下述范围完成。E5 包含动态数组、基础透视表、LET/静态引用/五个统计函数、Strict OOXML 及线程评论原位保留；每项限制仍独立适用，不代表完整 Excel 兼容。最终集成与本地包证据见文末。源码修改尚未提交 Git。

## 实施顺序与验收

1. XC-0501：先建立有界二维数组结果、溢出所有权及统一分配，再实现 SEQUENCE/FILTER/SORT/UNIQUE、查找数组返回和溢出引用。覆盖阻塞、合并/Table、旧结果清理、依赖、取消、编辑/撤销、结构变换及 XLSX/XLSQ 往返。未建模的旧 CSE/外部数组仍保持受控保留。
2. XC-0502：建立透视定义/缓存与过期状态，支持本地数据来源、行/列字段、求和/计数、创建和显式刷新；验证已有文件读取、编辑后保存重开、源变化提示与原子撤销。切片器和数据模型单列。
3. XC-0503：按可验证的小项落实复杂公式（优先 LET、OFFSET/INDIRECT 与明确的统计函数），逐项记录绑定、依赖、重算和导出边界。LAMBDA、旧 CSE 和复杂引用不得因同属 E5 而被误标支持。
4. XC-0504：将扩展文件变体和对象拆为实际实现项，优先 Strict OOXML 读写兼容和可安全保留的线程评论/外部链接；未知关联继续报告，外部内容不自动访问或刷新。调查本身不计实现完成。

各阶段只跑受影响的专项测试；独立 ZIP/XML 输入、真实 Excel 对照、Avalonia 路由及发布自检分别记录。历史 WPF 检查不作为当前 UI 验收。实现和证据均留在 QuickExcel；不改同级 ExcelIO。

## XC-0501 基础切片历史

- 新增不可变二维公式值，形状和元素数量独立校验。初始预算：单数组最多 100,000 元素、文本总量 4 Mi 字符；沿用单格文本和工作表坐标限制。此预算不等于整个工作簿的容量。
- 文件格式、查找数组返回、完整结构操作和产品验收尚在实施，暂不宣称 XC-0501 完成；未发布本轮包。
- 已实现 SEQUENCE/FILTER/SORT/UNIQUE 数组求值、二维广播、逐元素错误和聚合消费，已加入生产计算注册表。统一溢出分配与 DocumentSession 原子发布已接入，子格保留空白输入及显式所有者，结果缩小清理旧子格，输入修改/撤销保留缓存状态。
- `DynamicArrayTests|DocumentSessionTests|CalculationAdapterTests`：59 项通过，证据 `artifacts/e5/foundation/e5-array-session-01.trx`。排序取消曾被 LINQ 包装为比较异常，已改用可取消的稳定归并排序；失败证据保留。
- `scripts/test-excel-dynamic-arrays.ps1` 生成 36 条独立 Excel 动态数组公式、864 个输出位置。Excel 16.0 build 20527，语言 2052，重算/保存重开稳定；原始文件与报告在 `artifacts/e5/excel-arrays-01/`。该脚本复用 scalar 驱动，驱动输出的“864 scalar formulas”实际是 864 个观测格，不是公式数量。
- Excel 对照发现 SEQUENCE 零维输出应为 #CALC!，已修正；空白排序始终在末尾，输出时转为零，UNIQUE 判等前仍区分真正空白、数值零、空文本和文本零。独立预期保存在 `Fixtures/excel-dynamic-arrays.json`，不由产品求值器生成。
- 新增 `#` 锚点溢出引用、迭代依赖规划和当前轮派生格读取。普通标量工作簿仍走原增量/批处理路径；数组工作簿在共享五秒截止时间内最多 32 次区域规划，只有稳定结果才能提交。已覆盖消费者先于生产者、旧区域缩小、子格循环和占位冲突。`e5-array-round-02.trx`：13 项通过。
- 新增四个函数的 24 条真实 Excel 参数数量边界，原 67 函数证据保持不变。`e5-related-01.trx` 中 121 项通过，唯一失败为注册表新增项尚未补证据；补齐独立 Excel 证据后 `e5-arity-02.trx` 的 5 项边界测试通过。最终数组/会话/原计算器组合记录 `e5-array-excel-02.trx` 为 60 项通过（新增调度测试在之后单独执行）。
- 临时保存保护：原生动态数组 metadata 接入前，公式保存拒绝含动态数组的新模型，避免把派生子格写成普通输入；明确的值快照导出保持可用。后续完成原生往返及独立校验后移除此临时保护。
- 本切片末尾检查：`DynamicArrayTests|CalculationBoundaryTests` 18 项通过，`artifacts/e5/foundation/e5-foundation-final.trx`。未跑全量回归、未发布、未提交 Git。E5 目标仍处于 active，继续执行上述剩余工作。

### 动态数组验收完成（2026-11-03）

- 移除上述临时保存阻断。实现 XLDAPR、cellMetadata、原生 `_xlfn`/`_xlws` 和 ANCHORARRAY 编码；与富错误元数据共存，#SPILL! 保存实际阻塞尺寸。XLSQ v2 声明 `dynamic-arrays` 必需特征；SheetId 迁移后子格所有者重新绑定。
- 原生元数据只接受已校验的绑定、形状与完整缓存。旧 CSE/未知语义保留只读原包；伪造所有者、重叠、子格独立公式和超预算缓存拒绝。动态数组改写保存要求当前修订的新结果；不把旧缓存写成新结果。后写校验同时检查数组形状、派生值、所有权和阻塞尺寸。
- 结构命令保留显式数组模式，整体移动/插入前置行列/复制工作表后重建派生区域；部分剪切及不能保持所有权的局部结构修改原子拒绝。整块复制仅复制锚点输入，局部复制派生值；缩小结果保留显式子格样式。转置公式尚未支持，允许明确转置值。子格不能独立编辑或清空；修改/清空锚点、样式、撤销均通过会话。
- Avalonia 显示溢出边框，子格公式栏显示所属锚点且只读；保护隐藏的锚点公式不会由子格泄露。锚点仍可正常编辑。检查了实际网格截图，不将自动路由验收等同物理 IME/设备门禁。
- 数组规划按统一 5 秒期限、32 次区域迭代执行，循环与不稳定计划不提交半轮结果；不存在以旧派生值替代本轮结果。脚本在纯数组规划稳定后按最终优先拓扑执行一次；暂不合批数组工作簿的脚本调用，标量路径原有合批保持。数组尺寸依赖脚本或不可用函数时整轮标记 Blocked 并保留旧缓存，避免猜测未知区域或重复调用有状态函数。
- Excel 对照现为 49 条独立公式、1,176 个观测格，含混合类型、空白、错误、IF/IFERROR/IFNA 和 XLOOKUP 多格返回。原始输入及 Excel 来源在 `artifacts/e5/excel-arrays-03/`；最终产品文件的真实 Excel 重算/重开对照在 `artifacts/e5/excel-array-output-03/`，全部一致。IFERROR/IFNA 较大备用数组的扩展规则由此修正，失败证据保留。
- `artifacts/e5/codec/e5-array-interaction-01.trx`：101 项计算、复制、结构、工作表及查找相关测试通过；`e5-array-codec-final-01.trx`：43 项数组/保留公式/参数边界通过；`e5-array-boundaries-01.trx`：18 项新增样式、Table/工作表边缘、名称/直接范围及存储测试通过。均为相关专项，不是全量回归。
- `artifacts/avalonia/workspace/20261103-142952-182de834/`：Release 构建零警告/错误，83 项网格专项通过（含新增动态数组路由）；`artifacts/e5/codec/array-schema-03.json`：Open XML 校验零问题。原生阻塞错误另由真实 Excel 重开校验，证据在 `artifacts/e5/excel-array-errors-output-01/`。
- 明确边界：单数组 100,000 元素/4 Mi 文本字符；总格数、历史内存及取消门禁同时生效。支持四个数组函数、算术/比较广播、聚合消费、三种分支和 XLOOKUP 单一查找值的行/列输出。未把多查找值、所有旧函数的数组提升、脚本决定形状、旧 CSE、LAMBDA 或任意新函数标记支持；导入未建模语义继续受控保留。

## XC-0502 基础透视表验收（2026-11-03）

- Core 新增有界不可变透视缓存、源定义与结果定义，缓存记录和显示格分开。支持本地 Table/连续区域来源、单行字段、可选单列字段、单个 SUM/COUNT 及行列总计；首行需唯一非空文本标题。创建、显式刷新、转换为普通结果值均经过 DocumentSession，可原子撤销。
- 字段类别保留类型；当前基本排序采用已有的序数、不区分大小写比较。未把任意手动排序、按本地语言排序或多层布局标记兼容。数字参与 SUM，非空文本参与 COUNT，空白度量不计数；真正空白结果保持空白。错误值、分组、多个行/列/度量、筛选字段、命名透视样式、切片器、OLAP/数据模型等未支持语义仍保留来源包，不猜测刷新。
- 源单元格输入或重算结果变化会保留旧透视结果并标记过期；源字段变化时拒绝猜测列对应关系。刷新需要源公式有当前修订的结果。普通保存保留过期缓存及 invalid 状态；值快照导出需明确接受旧结果，并去掉透视定义。缓存刷新不会自动访问外部数据。
- 输出区域不能覆盖源、已有内容、表、合并、其他透视或附件，动态数组也不能溢入透视区域；结果不能直接编辑。已支持的内容、格式、冻结/缩放/尺寸、打印和筛选操作保持统一命令路径；地址重排、源表结构重建暂不放开，可先转换为值。未建模透视定义的工作簿保留原结果并限制内容编辑，始终显示“可能过期/不能刷新”。
- 原生 `pivotTableDefinition`、`pivotCacheDefinition`、`pivotCacheRecords` 及工作簿/工作表关系已读写；缓存支持索引共享项及内联类型值，校验集合数量、索引、字段角色、输出布局、值和范围。未知扩展不静默丢弃。XLSQ v2 增加 `worksheet-pivots` 必需特征。单缓存最多 50,000 行/128 字段/200,000 值/16 MiB，每表最多 16 个透视定义、32 MiB 缓存；输出最多 20,000 格，历史预算和文件读取预算另行生效。
- Avalonia“数据 → 透视表”提供创建（来源/字段/方式/位置选择）、刷新和转换为值。结果单元格只读，状态栏持续显示待刷新数量；操作重用版本门禁和撤销历史。`artifacts/avalonia/workspace/20261103-150607-29d72d4f/`：90 项网格/对话框专项通过，Release 构建零警告/错误；实际网格截图已检查。这不是物理设备门禁。
- 独立 Excel 两个透视样本（共享缓存，SUM 和 COUNT，含空白/文本/负数/零）及输入修改后的预期在 `artifacts/e5/excel-pivots-03/`，测试预期保存于 `Fixtures/excel-pivots.json`。早期脚本按 COM 枚举顺序比较导致误报，改按表名排序后重开稳定；失败证据保留。
- `artifacts/e5/pivot-native-import-02/`：真实 Excel 原文件的两个基础透视定义成功导入，来源修改后标记过期，刷新并另存，原文件哈希不变。原样式中的 pivotButton 等其他未建模格式属性仍按原兼容报告处理，本专项另存明确移除这类非透视定义的未支持属性，不把它们冒充已建模。对应结果由真实 Excel 再验于 `artifacts/e5/excel-pivot-output-02/`：刷新前、刷新后、保存重开均与独立 Edited 预期一致，源文件未变。
- 产品新建输出也由 Excel 实际打开/刷新/重开验过，见 `artifacts/e5/excel-pivot-output-01/`，与独立 Initial 结果一致。`artifacts/e5/pivots/pivot-schema-01.json` 和 `pivot-schema-02.json` 均为零 Open XML 校验问题。
- 独立内联 ZIP/XML 夹具覆盖旧缓存显示、读取后编辑/刷新、XLSX/XLSQ、外部源/不支持汇总保留、坏共享索引/记录数拒绝、值导出和源文件保持。`e5-pivot-codec-02.trx`、`e5-pivot-final-01.trx`：10 项通过；视图/尺寸补充断言最终记录 `e5-pivot-view-03.trx` 1 项通过。相关存储组合曾暴露字面量 #SPILL! 的默认 1×1 尺寸与空元数据比较问题，已修正默认表示；`e5-pivot-related-02.trx` 19 项通过，先前 125 项通过/2 项失败证据保留。没有运行全量回归。

格式关系依据：[Microsoft Open XML 透视表定义、缓存和显示格](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables)。本实现通过独立 XML 与真实 Excel 双向验证，不仅依赖自身写入器。

## XC-0503 复杂公式子项验收（2026-11-03）

- LET 支持顺序绑定、嵌套/遮蔽、常量/引用/数组类型、按需求值和单次记忆。局部名称不泄漏到定义名称解析；复制/移动/改名遍历识别词法作用域。把全局名称改为已有 LET 局部名会明确拒绝。绑定后深度及局部引用链限制为 128，保留既有 8,192 字符/2,048 语法节点/16,384 绑定节点预算。依赖保守包含声明的引用，脚本和最终计算仍使用现有受限执行路径。
- 独立 Excel 文件揭示 LET 参数必须用 `_xlpm.` 编码；编码/解码只修改词法参数，不改字符串、表名、列名或工作表名，保留原格式。首轮内部往返通过但 Excel 拒绝打开的失败证据已保留，修复后实际 Excel 普通打开、重算、保存重开均通过。
- OFFSET 支持本地矩形基准及常量行/列偏移、尺寸（含负尺寸、截断小数、显式缺省）；INDIRECT 支持常量 A1/R1C1 单格/矩形地址及相对 R1C1。目标在绑定时确定，进入精确依赖索引；可移植公式仍保留原函数文本，INDIRECT 字符串不会因结构移动被错误重写。产品坐标外、动态值决定目标、名称地址、未建模整行/R1C1 整列、外部/结构化文本等保持保留边界，不访问外部目标。缺少现代数组元数据的旧多格 OFFSET/INDIRECT 不猜测隐式交叉语义。
- 新增 VAR.S、VAR.P、STDEV.S、STDEV.P、MEDIAN；支持引用与直接常量各自的文本/布尔/空白规则，传播错误，样本数不足返回相应 Excel 错误。方差采用偏移和补偿累加，避免大基数抵消。单调用最多 100,000 个数值，沿用计算取消/时限；方差/标准差最多 254 个参数，中位数最多 255。
- `Fixtures/complex-function-spec.json` / `excel-complex-functions.json`：56 条独立 Excel 公式，初始/修改值及保存重开稳定。39 条新增参数边界见 `excel-complex-arity.json`；原始证据 `artifacts/e5/excel-complex-04/` 和 `excel-complex-arity-03/`。LET 同名重复参数由 Excel 拒绝，不声明支持。
- `artifacts/e5/complex-native-03/`：真实 Excel 输入的计算、增量编辑、撤销/重做及 XLSX/XLSQ 重开，共 448 项检查和 1 个超出产品坐标的原包保留检查全部通过；写出原生包的 Open XML 校验零问题。`complex-native-excel-initial-01/`、`complex-native-excel-edited-01/` 由真实 Excel 验证初始/修改输出，最终 `complex-native-verdict-01.json` 零失败且源哈希不变。产品新建路径另有 `excel-complex-output-03/` 验证。
- `artifacts/e5/complex/e5-complex-related-05.trx`：116 项相关测试通过，含数组/日期/文本、结构/名称操作；内联独立 ZIP/XML 验证未建模引用保留旧缓存、只读与原包逐字节另存。真实 JSC 测试验证 LET 重用有状态脚本表达式只执行一次，未使用变量不执行。
- 逐项边界：LAMBDA/递归及闭包尚无执行模型；动态 OFFSET/INDIRECT 需要新的运行时依赖计划；旧 CSE 需要独立编辑/隐式交叉契约；3D、联合和交叉引用需要多区域引用值及结构规则；其余财务/统计函数无本轮独立样本。这些仍保留/限制，未因 XC-0503 子项验收而标记支持。依据包括微软 [LET](https://support.microsoft.com/en-gb/excel/functions/let-function)、[OFFSET](https://support.microsoft.com/en-us/excel/functions/offset-function) 和统计函数说明，以及上述本地真实 Excel 样本。

规范依据：[Microsoft 动态数组及溢出行为](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior)。实际支持范围以实现及下方验收证据为准。

文件结构依据：[动态数组属性](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d1ef676d-d970-4ddd-b642-62f5e0c291e0)、[溢出富错误类型](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7a27b5fb-48a8-4a1c-b67a-2ca5edc07514)。#SPILL! 的 errorType 为 8，并含 colOffset/rwOffset；不能凭普通错误字符串宣称原生写回完成。

## XC-0504 扩展文件与对象验收（2026-11-04）

### Strict OOXML 读写子项

- 以包入口关系和工作簿根命名空间共同识别 Strict，矛盾输入拒绝。仅对明确的 SpreadsheetML、关系、DrawingML 及属性命名空间建立适配；用户文本、公式字符串和外部 Target 不做替换。原始包字节保留，解析副本适配现有模型，编辑另存继续使用 Strict 命名空间、关系类型及 `conformance="strict"`，并非一律转换为 Transitional。
- 支持 Excel 原生 `t="d"` 本地 ISO 日期/时间，映射 1900/1904 序列及虚构的 1900-02-29；不把时区或未知 ISO 形式猜成数值。已知 DrawingML 百分比按明确字段转换，精度/范围不满足时拒绝。主题、单元格及公式继续采用各自既有白名单；未知对象保留/报告，编辑后不能静默删除。
- XLSQ v2 声明 `strict-ooxml` 必需特征，重开仍保持内部 XLSX 变体。独立 ZIP/XML 测试覆盖日期系统、用户文本中同名 URI 不变、XLSX/XLSQ 修改往返、未知对象原字节另存及损失阻断。
- 真实 Excel 16.0 build 20527 通过 `FileFormat=61` 生成独立输入：`artifacts/e5/excel-strict-types-01/` 有 10 个日期/文本/转义/布尔案例，`excel-strict-complex-01/` 有 56 个复杂公式案例。产品原生读取、计算、编辑、撤销/重做、两格式重开分别为 80 和 448 项检查；另有 1 个超产品坐标的原包保护检查。结果在 `strict-types-native-03/` 与 `strict-complex-native-01/`。
- 产品写出文件由真实 Excel 普通打开、重算和保存重开复验，初始/修改结果均与独立预期一致；最终零差异判定见 `artifacts/e5/extended/strict-types-verdict-01.json`、`strict-complex-verdict-01.json`。原输入 SHA-256 未变。
- Open XML SDK 会翻译 Strict 关系/命名空间，但本机版本对 Excel 原生的部分百分比文本仍报 Transitional 整数类型错误。独立审计因此记录原包事实/哈希，仅在验证副本中投影已知百分比字段；上述输出在此明确配置下零校验问题。这不是完整 ISO Strict 校验认证，真实 Excel 重开为另外一层证据。保留最初的原始校验失败记录。
- 已验收范围为上述单元格、公式、样式/主题和已有元数据路径；未把所有 Strict 绘图、嵌入对象及任意扩展声称为已验证。

### 线程评论原位保留子项

- 验证本地 threadedComments/person 部件、GUID、作者和根评论/回复同格关联、时间、文本及旧版显示副本。支持基本文本讨论、回复和完成状态的透明保留；不查询账号、不联网、不发送通知，也不提供讨论编辑或云协作。
- 值/公式/格式的受支持原位修改可正常保存。线程、作者和 Transitional 旧版批注副本原始部件逐字节带回；XLSQ v2 声明 `preserved-threaded-comments`。兼容报告明确显示“仅保留”，修改值后仍不冒充损失。既有批注角标/悬停可显示来源提供的旧版副本，来源没有副本时不合成讨论 UI。
- 保留描述以工作表名和单元格绑定；结构重排、工作表身份/顺序改变、复制含讨论的区域和编辑旧版批注副本明确拒绝。普通批注及 note-only VML 的保护不放松；值快照和可移植公式导出也保留已验证的讨论附件。
- 未知 mentions/扩展、额外部件关系仍进入原包保护/损失报告；不访问任何外部关系。悬空作者、重复标识、无效回复关联等损坏输入拒绝。预算为每部件最多 4,096 条评论、4 MiB 文本，每个作者部件最多 4,096 人，工作簿最多 4,096 个讨论锚点；全包读写预算另行生效。
- `artifacts/e5/excel-threads-01/`：真实 Excel 创建两个根评论和一条回复，保存重开稳定。`threads-native-01/`：修改 A1 后两个锚点保留，三个源附件部件逐字节一致，输出 Open XML 校验零问题，源哈希未变。`threads-excel-output-01/`：产品另存结果经真实 Excel 打开/重开，两个讨论及回复均与独立源一致。作者等来源数据只在本地 artifacts，提交的夹具使用虚构固定身份。
- 独立 ZIP/XML 覆盖 XLSX/XLSQ 修改、撤销/重做、值/可移植导出、结构/复制/批注保护、坏身份拒绝，以及未知内容保留；与 Strict 合计最初 13 项通过，见 `artifacts/e5/extended/e5-extended-01.trx`，最终导出组合见下节。

关系与格式依据：[Microsoft 线程评论部件](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8)、[作者部件](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1a170d26-42a2-46f0-b2b6-0ff1dec1c344)、[SDK Strict 支持说明](https://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk)。

### 本项明确保留的范围

外部工作簿/数据链接仍采用既有原包、只读公式或明确旧缓存值导出路径，不进行自动访问、刷新或关系重建。本轮没有新增外部链接可编辑模型。更多复杂绘图、嵌入对象、完整线程编辑/提及和 Strict 全对象验证均未完成；这类输入继续报告/保护，不能以“E5 完成”覆盖各特性的 R/D/C/E/W 限制。

## E5 最终集成（2026-11-04）

- 解决方案 Release 构建零警告/错误：`artifacts/e5/final/release-build.log`。
- 764 项受影响的计算、存储、结构/剪贴板、会话及元数据测试全部通过，零跳过：`artifacts/e5/final/e5-integration-01.trx`。没有运行全量 worker/安全/平台回归。
- 导出审查补齐显式数组模式，并将非可移植数组的派生子格和引用子格的消费者一并冻结；旧缓存需明确选择，不能保留孤立的溢出所有权。40 项导出/Strict/线程/诊断组合通过：`e5-export-final-02.trx`。首轮测试误把替换公式后已清除的缓存当成可导出值，改为独立旧缓存夹具后验证正确，失败记录保留。
- 当前 Avalonia 90 项网格/透视对话框/数组交互检查通过：`artifacts/avalonia/workspace/20261103-160533-f9d21b95/`；已检查实际网格截图。此路径时间为 UTC，验收日期为本地 11 月 4 日。
- M4 严格策略、M5/M6 实机 IME/DPI/辅助功能/打印/干净设备等门禁仍独立待验；JSC-MEM-001 按用户决定延期。本轮没有修改隔离权限，也不宣称 Beta 验收完成。
- 使用 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke -CreateArchive` 生成并激活新的本地 Windows 自包含包。`SkipTests` 跳过与本次修改无关的发布脚本安全/平台组合，前述 764 项专项已独立完成；生产自检仍完整执行。八项自检包括 typed 保存重开、实际隔离 JSC/TS worker、Monaco/PDF.js、PDF 输出和资源清理，全部通过。包证据：`artifacts/avalonia/releases/20261103-160612-6d18c466/evidence/`；生产编译有一条已有诊断关闭构建的 `Program.WebScriptsOnly` 未赋值警告，零错误。首轮命令因相对日志路径随脚本目录变化而失败，改用绝对日志路径后成功，未改动发布逻辑。
- 新包：`artifacts/avalonia/releases/20261103-160612-6d18c466/application/`，版本 `0.6.0-preview.8`；默认 `Start-QuickExcel.cmd` 通过 `artifacts/avalonia/active-windows.json` 指向该包，旧包保留。归档 `artifacts/avalonia/releases/QuickExcel-Avalonia-0.6.0-preview.8-win-x64-20261103-160612-6d18c466.zip`，SHA-256 `F434B0DFEB6F2819D468476850441E085E37792B9296B036C0D7EE8C8FDF86CB`。这是本机验证的预览包，不是外部发布或干净机器认证。
- 独立复核激活指针与 release.json 哈希、ZIP 中 736 个清单文件的长度/SHA-256 及无额外文件，全部一致；ZIP 为 228,905,311 字节，记录在 `artifacts/e5/final/release-verification.json`。最终源码身份、未提交 patch 和新增源文件副本由 `scripts/record-baseline.ps1` 采集至 `artifacts/e5/final/baseline-capture-*`；不以本机未提交源码冒充一个 Git 提交。
