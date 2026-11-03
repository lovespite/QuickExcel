# XC-0106：不支持公式的受控读取与保留

2026-11-03：XC-0501 已接入有界动态数组模型和原生 XLDAPR 元数据。已支持的四函数、XLOOKUP 多格输出、数组运算/分支及溢出引用可计算、编辑、往返；范围与证据见 [74](74-e5-implementation-log.md)。下表中动态数组的原包只读策略现仅适用于未支持的语义、旧 CSE、未知元数据或不完整缓存，不再覆盖所有动态数组。

2026-10-30 完成，提交为本记录所在的 `feat: preserve unsupported formulas with read-only safeguards (XC-0106)`。本项交付来源公式/缓存保留和编辑限制；数组计算、溢出所有权及结构化引用计算仍分别属于 E5/E3。

## 契约和范围

| 类型 | R / D | C | E / W |
| --- | --- | --- | --- |
| 可解析的未知普通函数、`_xlfn.` / `_xlws.` 前缀 | 原公式和来源缓存；报告函数与位置 | 不可用函数及已知传递依赖保留缓存并标 Blocked，独立公式仍计算 | 值/样式可编辑；导入工作簿的结构变换和排序拒绝。普通保存保留源码/缓存及 Excel 重算标记；其他未知部件仍按损失策略处理 |
| 传统数组、动态数组、结构化引用、外部引用、不能展开的共享公式、其他不支持语法 | 整簿只读，显示来源缓存；分类、位置和结果范围可查 | 全簿公式保留缓存并标 Blocked，不运行脚本 | 只可复制完整原 XLSX 包或存入 XLSQ；编辑、结构操作、删减特性的公式副本及可移植公式导出拒绝 |
| 不支持的自定义名称、名称绑定失败 | 同上；不能把未建模名称变成新鲜 `#NAME?` / `#VALUE!` | 同上 | 同上 |

只读策略刻意覆盖整簿，因为数组结果区域和不支持语法的依赖尚不能完整分析。数组从属格虽然投影为普通值，标题、兼容报告及状态栏明确其来源缓存身份。不能安全展开的共享从属格显示主表达式，精确 XML/引用属性留在原包，不能据投影编辑。数据表 `t=dataTable` 等其他不支持的公式类型仍按既有读取拒绝处理；容量、XML/ZIP 预算不放宽。

`DocumentSession` 拦截全部修改命令、脚本启用及伪造的新计算结果；拒绝不创建 dirty/撤销记录。导航允许且不标 dirty，保存时恢复源视图，不把只读导航写入原包。存储层再次比较内容，直接构造修改快照也不能绕过保护。桌面公式栏/脚本栏只读、网格编辑和粘贴停用，关闭只读文件后普通编辑恢复。实际重算入口显示“未重算，已保留来源缓存”。

XLSQ 仍是 v2；含保留公式时声明 `preserved-formulas` 必需能力，旧读取器明确拒绝。标准内嵌 XLSX 是唯一权威来源，重新读取时重建分类。只读文件不应用旧 JSON 视图覆盖，保证原包另存语义；无此类公式的旧 v1/v2 视图回退不变。

显式值快照导出需要接受旧缓存和已报告损失；数组结果区域有缺失缓存或完整性检查超过预算时，即使同意也拒绝导出。原包另存不受此限制。启用脚本只解除实际已注册函数的阻塞；未注册 Excel 函数不会被误当成成功的脚本结果，版本前缀也不交由脚本覆盖。

## 验收证据

真实 Excel 16.0 build 20527，64 位，语言 2052，自动计算 -4105。自建 artifacts 文件，隐藏的独立 COM 实例，普通打开、不请求修复；没有访问用户现有工作簿。来源和输出均全量重算、保存、关闭重开。

| 检查 | 结果 / `artifacts/compatibility` 下证据 |
| --- | --- |
| 公式保护、计算/名称/结构、typed 存储及原生脚本专项 | 159/159，`xc0106-final-related`；包括缺失数组缓存、伪造计算结果、部分循环、跨表依赖及注册/未注册 UDF |
| 独立比较工具专项 | 36/36，`xc0106-probe-final`；数组夹具未编辑时原包相同，修改路径明确拒绝 |
| 桌面实际处理器 | 84/84，`xc0106-wpf-final`（本项增加 10）；含来源缓存截图、公式栏、兼容对话框、保存和恢复编辑。非物理输入/OS IME 验收 |
| Excel 只读来源 | `xc0106-excel-readonly-03`：四表，传统数组、Table 结构化引用、SEQUENCE 溢出及 XLOOKUP 版本函数；52 个 UsedRange 格记录 Value/Formula/Formula2/Text/数组范围 |
| QuickExcel 只读往返 | `xc0106-readonly-final/formulas-report.json`：8 个公式缓存保持 Blocked，编辑被拒绝，原包字节相同、来源哈希不变 |
| Excel 读取只读输出 | `xc0106-excel-readonly-reopen-01`：52 格记录与来源相同，保存重开相同；最终输出与该已验收输出字节相同 |
| 未知普通函数对照 | `xc0106-excel-unknown-01`、`xc0106-unknown-final`：A1 从 9 改为 16，SQRT(A1) 保持旧缓存 3，依赖保持 4，独立 SUM(A1,2) 新结果 18。保存保留公式、源码缓存均可核对 |
| Excel 重算未知函数输出 | `xc0106-excel-unknown-final`：Excel 结果为 4、5、18，与 Excel 自身编辑参考一致，保存重开一致 |

Open XML 独立校验：未知函数输出 0 问题；只读源文件本身有 4 条 calcChain `c@s` 被校验器关联为 cellXfs 索引的问题，输出整包相同、无新增问题。报告同时保留源问题和输出问题，不把该文件宣称为 schema 全通过，也不删除 calcChain 来掩盖。Excel 普通打开、重算、重开正常。

初期 COM 脚本的字符串/数值动态赋值错误及测试断言失败保留在各 `-01/-02` 目录，不列作通过。最终 Desktop Release 构建通过，仅既有 NU1900 漏洞源不可达警告。本项未运行全量回归、发布包装或性能门禁。

复现入口：`scripts/test-excel-formulas.ps1 -OutputDirectory <新 artifacts 目录>`；未知函数源增加 `-Scenario Unknown`。然后 `dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- formulas <来源.xlsx> <新目录>`。Excel 输出复验传入 `-WorkbookPath <输出.xlsx> -ReferencePath <来源 excel-formulas.json>`，未知函数对照增加 `-ReferenceState AfterEdit`。

未知函数源仍有修订/拼音等未建模内容，首次编辑输出使用明确的兼容副本；本项不承诺任意模板小改无损。下一项 XC-0107 补富文本、超链接和传统批注。
