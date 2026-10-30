# M3 跨表与增量计算 / 0.4 计算预览

2026-10-30 状态说明：本页保留 M3/0.4 计算子集与当时验收；后续已接入独立 AppContainer worker、模块/actions、完整工作区入口和签名信任，不能从下文历史“尚未实现”推导当前状态。新增 Excel 公式、结构化引用和导出扩展按 [专项计划](26-excel-compatibility-roadmap.md) 推进，总进度见 [实施计划](05-implementation-plan.md)。现有公式子集的语义限制继续适用，除非后续实现记录明确更新。

日期：2026-10-27。基于独立 QuickExcel 仓库的 M2 提交 `813e91d`；本轮修改尚未提交。同级 ExcelIO 源码未修改。0.4 交付以下明确支持的计算子集，M4 worker、M5 完整结构交互和 M6 Beta 门禁继续后续推进。

## 任务与实现

| 任务 | 本次结果 | 验证入口 |
| --- | --- | --- |
| CALC-001 | Core 中的有界 A1 AST、表名/名称绑定；Infrastructure 类型化求值，桌面已切换 | M3CalculationTests |
| CALC-002 | 工作簿稀疏矩形引用索引、完整重算、迭代 SCC；循环成员和循环下游诊断分开 | 跨表链、4000 格依赖链、循环夹具 |
| CALC-003 | 比较不可变输入得到编辑集，计算新依赖闭包；确定性结果复用并重新标记当前 stamp | 100 轮固定种子编辑/公式替换/清空/撤销对照全量；2 重算 + 1 复用断言 |
| CALC-004 | 行列插入/删除/块移动、同表/跨表剪切移动、表改名；名称一起更新，一次撤销 | M3StructureTests |
| CALC-005 | Fresh/Error/Blocked/旧缓存分离；带工作表和地址的诊断；继续使用四元组提交门禁 | 错误类型、禁用脚本、取消和旧结果拒绝 |
| JS-005/006 | v1 扁平范围保持；v2 描述符、JS 命名空间、参数/返回类型检查、显式二维范围 | M3ScriptTests，真实 QuickJS |
| JS-010 M3 契约 | 已授权脚本每轮新上下文、按表序/行列序稳定拓扑重算；无 UDF 复用 | 有状态 NEXT 闭包连续三轮，独立标准式也重算 |
| FMT-007 | 保存缓存聚合状态；重开仍视为 Stale；可移植公式导出冻结 JS/不支持/循环种子及跨表依赖 | M3StorageTests、WPF calculation smoke |

图和输入差异每轮仍重建/扫描；增量优化的是求值闭包，不宣称依赖图更新已完全增量化。初次计算、强制全量、会话/表结构/名称/日期系统/脚本代次变化会完整重算。TODAY/NOW 每轮使自身及下游失效。包含未知函数的未授权工作簿也采用保守全量计划。

计算结果来自计算器内部的完整成功轮，绝不从文件的 Fresh 标记直接建立复用缓存。取消/超时不发布部分结果；复用结果在发布时统一携带当前修订与代次。输入改变至计算完成期间，会话保守地把已有公式缓存标为旧值。引擎不会向 ExcelIO 全局 cell 事件注册回调。

## 公式支持矩阵

2026-10-30 增补：XC-0205 已实现七个条件统计函数及实际 Excel 对照，参数模式、类型规则和数组条件保护见 [44](44-conditional-functions.md)；XC-0206 查找函数边界、现代标量模式与 XLSX 别名见 [45](45-lookup-functions.md)。下表保留原 M3 基础子集。

| 范围 | 支持与限制 |
| --- | --- |
| 标量 | Number/Text/Boolean/Blank/Error 保持身份；数字外观文本不自动变成数字；错误外观文本不自动变成 Error |
| 语法 | 科学记数法、双引号转义、比较/连接/四则/幂/百分号、括号；完整消费 token，无法解释的输入产生错误 |
| 引用 | A1、绝对/相对/混合、矩形、A:A、跨表、中文及转义单引号表名；旧裸列 A/AB 在没有同名定义时继续表示整列 |
| 名称 | 工作簿/工作表作用域，局部优先；常量、公式和别名；含单元格的定义须带表名并使用绝对引用；通过 XLSX/XLSQ 往返 |
| 聚合 | SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/PRODUCT；引用和范围中的文本/布尔不计为数字，COUNTA 计入空文本；直接标量参数按函数规则显式转换 |
| 条件 | IF、IFERROR 懒求值；AND/OR/NOT；图仍为静态全分支依赖，因此未选分支中的循环或禁用脚本依赖也可能使公式受阻 |
| 数学/文本 | ROUND（中点远离零）、ABS；CONCATENATE/LEFT/RIGHT/MID/LEN/TRIM/UPPER/LOWER；TRIM 折叠普通空格 |
| 查找 | INDEX 标量索引、MATCH 一维、VLOOKUP；精确匹配以及要求已排序输入的近似模式；不支持通配符匹配、动态数组或 INDEX 返回引用 |
| 日期 | TODAY/NOW 返回当前本地时间对应的 1900/1904 数值序列，显示由既有样式决定；不隐式创建 JS Date |
| 不支持 | 外部工作簿、3D/联合/交叉引用、整行范围、数组/溢出、结构化表引用、动态引用函数与未列出的函数；不承诺完整 Excel 一致性 |

运算符顺序对照 [Microsoft 的 Excel 运算符说明](https://support.microsoft.com/en-us/excel/calculation-operators-and-precedence-in-excel)：负号高于百分号和幂，同级从左到右。固定回归包括 `=-2^2` → 4、`=2^3^2` → 64。这里只引用运算符规则，不将其作为本实现的整体 Excel 兼容认证。

旧 `CalculationAdapterTests` 保留并明确属于 LegacyWorkbookCalculator。跨表 Blocked、脚本覆盖内置、旧聚合/转换等历史行为不迁入新引擎；类型保持、错误传播、惰性求值、稀疏范围、运行时释放等契约由 M3 固定预期夹具重新验证。独立 ExcelIO.Formula 项目既有测试不等同于新求值器验收。

## 结构编辑契约

`ShiftAxisCommand` 支持行列插删；`MoveAxisCommand` 的 destination 是移除块后的最终索引；`MoveCellsCommand` 是剪切移动，目标原有内容被覆盖，源范围中的空白也会清空对应目标。区域和公式来自同一不可变快照，重叠移动不会逐格覆盖读取中的源数据。

结构变化移动引用目标，不因 $ 标志而固定旧坐标。插入扩大内部范围，删除缩小范围，全部被删的引用成为 #REF!；跨表移动的公式保留原引用目标。表改名使用转义名称并保留 SheetId。每项命令先完整准备，成功后只生成一条历史；撤销/重做恢复公式、名称、位置与元数据，Revision 继续递增。

会导致引用区域不连续的移动、仅覆盖被引用范围一部分的区域移动，以及带图片/合并单元格的结构移动会原子拒绝。无法解析的公式会阻止结构变换。插入越界也拒绝，不静默丢弃末尾数据。当前没有工作表删除/复制命令或完整结构编辑 UI；这些与图片/合并锚点交互继续 M5。

## 脚本版本与执行

旧的 `register("DOUBLE", n => n * 2)` 和一维行主序范围继续有效，v1 的 Blank 参数仍为 ""。v2 在函数描述符显式选择：

```javascript
quickExcel.functions.register({
    name: "JS.TOTAL", apiVersion: 2,
    parameters: [{ name: "cells", type: "range" }],
    returns: "number", profile: "conservative"
}, cells => cells.values.reduce((sum, value) =>
    sum + (typeof value === "number" ? value : 0), 0));
// =JS.TOTAL('数据表'!A1:B4)
```

v2 范围是冻结的 `{kind:"range", rows, columns, values}`；values 是行主序标量数组，Blank 为 null，空文本为 ""。标量参数同样保留类型。错误在进入 JS 前传播，不把任意对象识别为宿主错误。描述符参数数量固定，支持 any/number/text/boolean/blank/range，返回类型为 any 或标量类型。对象/Promise/非有限返回仍拒绝。

描述符 `apiVersion:2` 在本阶段表示函数参数桥版本；XLSQ 的 `scriptApiVersion:1` 仍标识现有单模块加载 API。函数的显式选择写在源码中，保存不执行源码来推断版本。旧应用可以保留源码但不支持执行对象描述符；不要用旧应用计算并覆写带新公式的工作簿。本次没有启用文件级模块 API v2。

与 0.3 的差异：新计算器禁止脚本覆盖 SUM/IF 等内置名称，冲突导致整轮不提交，提示改名为 JS 命名空间。仅支持 conservative；`profile:"strict"` 明确拒绝，声明无法证明任意 JS 纯净。所有已启用脚本轮重建上下文并全量重算，不提供共享参数缓存。进程内执行、单调用原生截止时间和只能在调用间检查托管取消的限制保持，worker/OS 隔离仍属 M4。

## 缓存与导出

XLSQ 仍仅在 workbook.xlsx 保存公式及缓存值。document.json 的 calculation 保存引擎标识、输入/源码/绑定修订、源码集合摘要和 unknown/fresh/stale/partial 聚合状态；打开始终撤销授权，并将文件公式缓存视为 Stale，不能以落盘状态代替一次当前会话计算。

桌面有两个导出入口：

- **导出值快照**：所有公式变成类型化值，并移除定义名称；缺失缓存拒绝，旧缓存需要明确选择。
- **导出公式**：JS、循环、不支持表达式及它们的传递下游冻结为值，跨表和名称展开均参与依赖闭包；独立受支持公式保留，名称展开为实际引用。保留公式的旧/缺失缓存不写出，Excel 可重算；需要冻结但缺失结果时拒绝，冻结旧缓存需明确选择。

宿主 Cycle/Script 错误冻结时转换为标准 #VALUE!，原 XLSQ 保留宿主错误身份。导出不带脚本文件、不改变原文档或 SavedRevision；同目录临时文件重开验证后再原子替换。兼容性损失仍须明确选择另存受支持副本。

## 预算与验收

公式文本 8192 字符，语法最多 2048 个解析项、128 层；名称最多 4096 个，绑定深度最多 16、单次展开最多 16384 节点，展开后的树仍限 128 层。引用项和依赖边分别限 2,000,000；整轮 5 秒截止。JS 单范围最多 1,000,000 个位置，此外继续受 native 输入字节预算约束。聚合采用流式累计，不为整列空格创建单元格。

候选验证已完成：Core 328/328，native smoke、WPF 原有编辑/保存/关闭场景及新增 M3 场景全部通过，候选位于 `artifacts/m3-candidate`。之后补充负号/幂的 3 个固定预期回归，最终验证结果记录在下方。历史断言失败是 ImmutableArray 集合断言使用了容器相等比较，改为逐项数组比较后通过，没有删除失败场景。

证据目录：`artifacts/validation/release-0.4-preview`。WPF 新场景经过真实组合根/保存入口，验证跨表编辑、表改名后保存重开和可移植公式；未自动操作原生文件选择对话框，也不替代 M5 全部交互。已检查 WPF 生成的界面图，导出按钮正常显示。最终包仍需通过独立 ZIP 解压复验才计放行。

未执行干净机器/签名/安装器/多进程 worker/OS 隔离或 p95/RSS 基准；本次是计算预览，不使用完整 Beta 标签。

## 最终验收（2026-10-27）

| 检查 | 实际结果与证据 |
| --- | --- |
| Release solution | 0 警告、0 错误；artifacts/validation/release-0.4-preview/build-final.log |
| Core 全量回归 | 331 通过、0 失败、0 跳过（既有 266 + M3 65）；release-final.log；较早的 m3-core.trx 为 328 例候选结果 |
| 原生运行时 | PASS；10,000 次调用 170 ms，仅为单次 smoke 样本，不是 p95 性能承诺 |
| 发布版 WPF | 既有编辑/保存/8 类真实关闭、M2 类型迁移及新增 M3 跨表/改名/导出全部 PASS，退出 0；artifacts/smoke/20261027-073010-356fc4d4 |
| ZIP 独立复验 | 哈希及 sidecar 一致，新解压目录运行完整 WPF smoke，退出 0，四个结果文件为新 PASS；D:\Projects\QuickExcel\artifacts\validation\release-0.4-preview\extracted-smoke-8934cc6e93bf481d9790a14c6acd4a00 |
| 默认启动目录 | artifacts/QuickExcel 已更新，应用/Core/Infrastructure/Scripting/native 五个 DLL 与 ZIP 一致 |
| 源码边界 | 基于 813e91d，当前修改未提交；同级 ExcelIO tracked diff 为空 |

最终包：artifacts/releases/QuickExcel-0.4.0-preview.1-win-x64-20261027-073010-356fc4d4.zip，1246300 字节。
SHA-256：8CC3991FCFDF1F8F3D7A0F3FB8939A3D857B5C7FF9BD9EE4430F22CB645F98F9。

归档明细：artifacts/validation/release-0.4-preview/archive-validation.json。发布包内有 M2/M3 支持矩阵、发布说明和真实 BUILD-METADATA。当前文件的这段外部验收记录在验证 ZIP 后补写；不改写已验证 ZIP。最终源码状态由 record-baseline.ps1 另存哈希、patch 与新文件副本。

M3 本矩阵内的实现与验收完成；下一阶段 M4。结构 UI、受保护图片/合并的变换、不连续区域、完整 Excel 行为和 Beta 门禁仍按已声明范围后续推进。
