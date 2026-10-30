# XC-0002：Excel 独立编辑预期与往返判定

2026-10-30 完成 XC-0002，提交为本记录所在的 `test: verify edited workbooks against independent Excel operations (XC-0002)`。本项在 [37 号业务台账](37-excel-business-baseline.md) 上补齐五类八条路径的独立 Excel 预期、模型/包结构差异、实际 Excel 保存重开与自动判定。不是任意 OOXML 文件的通用等价证明；复杂 E1 对象仍使用各专题最小样本及期望。

## 路径和判定

六个自建业务工作簿分别从同一来源执行：不编辑、首表 A2=99、A2 切换粗体、第 2 行前插入一行、删除第 2 行、B 列前插入一列、删除 B 列、首表改名“改名后”。**Excel 先独立执行命令并重算，生成 48 份预期**，QuickExcel 不参与预期生成。原文件不覆盖，所有路径有 SHA-256、来源检查及独立输出。

新增命令：

- `edit-matrix`：核对 Oracle 与来源指纹；每条编辑执行一次撤销和重做，再实际重算；保存后重开对照。包括原值/类型、公式、显示文本、整格字体/字号/粗斜体；不将 Blocked 旧缓存当作重算结果。
- Excel `-EditMatrix -VerifyDirectory`：打开 QuickExcel 输出、全量重算、另存新文件、关闭重开；固定捕获每表 A1:E7，含空白，逐格记录字体颜色、数字格式、行高/列宽，另记可见性、打印区域/标题、纸张/方向/适页、页眉页脚。该步只捕获事实，原始不相等路径写入 `RawComparisonDifferences`，待最终判定。
- `verify-edit-matrix`：绑定模型报告与同一 Oracle，并用输出哈希证明 Excel 检查的是该次输出；检查所有捕获字段和保存重开/来源不变标志。除被支持解析器证明等价的公式表示外，其他字段不放宽。

公式中的冗余括号、可选表名引号、删除引用后的 `数据!#REF!` 与 `#REF!` 由 AST 规范形式比较；原公式全文仍留在两侧证据。绝对/相对标记、实际工作表、运算优先级、参数、文字常量不被忽略。解析不了的差异一律失败，不能用相同缓存值掩盖不同公式。本轮 **14 处公式表示差异**属于预期，**0 处未预期差异**。例如 `A1*(B1+2)` 与 `A1*B1+2` 必须失败。

`OpenXmlAudit` 继续记录所有部件哈希、XML/内容/公式/样式/关系/元数据差异，独立检查 SDK schema、关系引用/目标及样式索引。原始 XML 的排序、内联/共享字符串或样式表追加可能不同；原始差异列表不直接等同于损失。固定样本的语义判定由独立 Excel Oracle 和完整捕获字段负责，不对未捕获的任意扩展 XML 宣称自动等价。

来源现代 Excel 的修订/扩展、拼音等未建模内容仍触发普通保存保护：**6 条不编辑原包保存，42 条编辑普通保存被阻止后生成明确命名的 `*.diagnostic-copy.xlsx`**。报告保留每条拒绝及损失诊断，未关闭警告或声称 42 次普通无损保存成功。后续功能只按确切支持范围减少这些诊断。

## 比较抓出的产品修复

1. 插入普通行列原先只移动已有内容，新增空白没有继承邻近格式。现按 Excel 默认操作继承上方/左侧的单格样式和行高/列宽，但不复制值、公式、富文本 run 或隐藏状态；第一行/列前无前邻，使用默认/正交轴已有格式。容量先检查，失败不留下历史或部分修改。插入的内容与元数据可一次撤销。
2. Excel 删除跨表被引用地址后会保留 `工作表!#REF!` 写法。解析器现在将这一确切形式识别为引用错误，原公式文本仍原样保存，计算保持 #REF!；没有扩展到任意未知引用语法。

四个真实 Excel 边界另测行/列第 1 或第 2 位置各插入两条，来源前邻隐藏：第一位置使用默认字体/尺寸，第二位置继承 Arial 14 粗体及尺寸，新增轴始终可见。此证据纠正了初版“首轴继承下一条”和“复制隐藏状态”的假设；不是所有结构能力已完成，附件、只读公式、分页删除等既有拒绝仍保留。

## 证据

本机 Excel 16.0 build 20527、64 位、语言 2052、自动计算 -4105；宏禁用、不更新外链，独立隐藏 COM，只操作自建 artifacts。普通加载，不请求修复。路径均在 `artifacts/compatibility` 下。

| 验证 | 结果 / 证据 |
| --- | --- |
| Excel 独立八路径 Oracle | 48 份自算/保存重开一致，`xc0002-excel-oracle-final/excel-edit-matrix.json` |
| 模型、撤销重做、保存重开、结构 | 48/48，45144 属性断言；48 输出独立 schema 0 问题，来源全不变，`xc0002-edit-matrix-final/edit-matrix-report.json` |
| Excel 检查实际输出 | `xc0002-excel-verify-final/excel-edit-matrix.json`；逐格值/显示/字体/数字格式/尺寸及打印元数据，保存重开一致 |
| 最终自动判定 | `xc0002-verdict-verified.json`：48/48，14 条等价公式表示，0 意外差异；绑定三个报告哈希 |
| 结构/视图/打印及新修复专项 | 45/45，`xc0002-structure-tests`；首轴/隐藏边界修正后相关 32/32，`xc0002-insert-boundaries` |
| 公式/存储/只读保护专项 | 56/56，`xc0002-formula-tests` |
| 工具差异/破坏注入专项 | 24/24，`xc0002-probe-tests`；最终判定/基线专项 17/17，`xc0002-probe-final`；删除关系/格式、改变公式/优先级/表名/绝对标记/字符串、丢失字段及打印区域均能定位失败 |
| Excel 插入边界 | 4/4，`xc0002-insert-excel-final/excel-insert-formatting.json`，对应脚本 `test-excel-insert-formatting.ps1` |

最终矩阵的来源均无手动隐藏且插入位置为第二条，之后的首轴/隐藏边界修正不改变这 48 条路径的结果；边界修正由独立 Excel 与 32 项专项复验。工具重算状态守卫只是拒绝假成功，已有 48 路径均完成实际内置计算。

初轮 `xc0002-edit-matrix-01/-02` 的格式继承失败与限定错误引用解析失败、首轮原始 Excel 文本差异均保留，未删改旧失败证据。未运行全量回归、发布包装或新增 WPF 检查；本项未改 UI。仅既有 NU1900 漏洞源不可达警告。

## 复现

先按 37 号记录创建新的业务来源，再执行（所有目标必须不存在）：

```powershell
.\scripts\test-excel-business-baseline.ps1 -OutputDirectory artifacts/compatibility/new-oracle -InputDirectory artifacts/compatibility/new-business -ReferencePath artifacts/compatibility/new-business/excel-business-baseline.json -EditMatrix
dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- edit-matrix artifacts/compatibility/new-business artifacts/compatibility/new-oracle/excel-edit-matrix.json artifacts/compatibility/new-edits
.\scripts\test-excel-business-baseline.ps1 -OutputDirectory artifacts/compatibility/new-observed -InputDirectory artifacts/compatibility/new-business -ReferencePath artifacts/compatibility/new-oracle/excel-edit-matrix.json -EditMatrix -VerifyDirectory artifacts/compatibility/new-edits
dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- verify-edit-matrix artifacts/compatibility/new-oracle/excel-edit-matrix.json artifacts/compatibility/new-observed/excel-edit-matrix.json artifacts/compatibility/new-edits/edit-matrix-report.json artifacts/compatibility/new-verdict.json
.\scripts\test-excel-insert-formatting.ps1 -OutputDirectory artifacts/compatibility/new-insert-edges
```

前两种旧命令 `baseline` / `compare` 保持原有含义：合成样本生成或原始差异检查，不会因本项完成而自动升级为 Excel 验收。XC-0003 的诊断定位收尾独立进行，E0 尚不整体关闭。
