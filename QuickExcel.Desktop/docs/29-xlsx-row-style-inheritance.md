# XLSX 行列样式优先级

更新：2026-10-30。对应 [26 号计划](26-excel-compatibility-roadmap.md) 的 XC-0102 首切片，**已实现并通过本地专项验证，XC-0102 整体仍未完成**。范围限定为单元格、行、列之间的样式选择，以及 `row/@customFormat` 的读写；命名样式、主题字体和完整 XF 关系仍待后续实现。修改基于 `600d50804fdb1d92573b0d52b677d0392d491be2`，未创建提交或新发布包。

## 1. 契约

- Core 的 `RowMetadata` 单独记录是否应用行样式；样式索引 0 可以是有效的显式行覆盖，不能用非零来判断有无行格式。
- 有单元格记录时使用已解析的单元格样式。不存在单元格记录时：已启用的行格式 → 列格式 → 样式 0。读取未指定 `c/@s` 的单元格也使用相同的行/列规则。
- XLSX `customFormat=1/true` 才启用行样式；`0/false` 或缺省均不启用。缺省 `s` 为 0。禁用的非零样式索引保留在模型，但不会影响显示或录入，也不会在保存时被启用。
- 仅调整行高/隐藏不创建有效行格式。显式行样式 0 与未设置行格式分开保存；原包不变保存仍保持来源字节，重写/导出使用已有 XLSX 字段，不升级 XLSQ 元数据版本。
- Core 旧的 `new RowMetadata(..., StyleId: 非零)` 调用继续默认启用；默认 StyleId 0 不启用。导入器总是显式传入 `ApplyStyle`。后续改样式索引的命令必须同时设置应用标志。
- 清除内容保留有效格式；清除本来不存在的单元格不创建默认样式覆盖或新撤销记录。默认样式的显式单格覆盖在非默认行/列下不得被稀疏化掉。

依据：[ISO/IEC 29500-1:2012 §18.3.1.73](https://www.reginfo.gov/public/do/DownloadDocument?objectID=136723002) 的行属性定义说明 `s` 仅在 `customFormat` 启用时应用；[Microsoft Row API](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row?view=openxml-3.0.1) 对应同一 SpreadsheetML 元素。使用单独的行标志，不混用 `cellXfs` 的 `apply*`：后者在 Excel 中还涉及父命名样式更新传播，见 [Microsoft XF 实现说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/59922f8b-0edc-4e93-a822-9f22254aec46)。

## 2. 验证

新增独立 SpreadsheetML 夹具 `row-column-styles`：列日期格式、显式行样式 0、禁用和缺省 customFormat、启用非零行格式、缺省 s/显式 true、仅行高、单格覆盖与空白格。它是第 10 个合成样本，不是 Excel 生成文件。

| 验证 | 结果 | 证据 |
| --- | --- | --- |
| 既有存储/日期显示/工作区/快照/结构/会话测试 | 260 项通过 | `artifacts/compatibility/xc0102-row-core-tests/ASUS_LOVESPITE-PC_2026-10-30_13_04_50_net10.0.trx` |
| 新增 `StyleInheritanceTests` | 最终 5/5 | `artifacts/compatibility/xc0102-row-core-tests/ASUS_LOVESPITE-PC_2026-10-30_13_05_41_net10.0.trx` |
| 比较工具测试 | 34/34 | `artifacts/compatibility/xc0102-row-probe-tests/ASUS_LOVESPITE-PC_2026-10-30_13_04_11_net10.0.trx` |
| WPF XLSX 渲染专项 | 26/26 | `artifacts/compatibility/xc0102-row-wpf-20261030-130651-7ef2fdd9/` |
| WPF 输入/选择/粘贴/清除/撤销/保存专项 | 92/92 | `artifacts/compatibility/xc0102-row-interaction-20261030-130750-9776ff0b/` |
| Desktop Release 构建 | 通过 | `artifacts/build/bin/QuickExcel.Desktop/release/` |
| 更新后的 10 样本 × 8 路径基线 | 72 输出均无独立结构错误，8 数组导入拒绝，17 普通保存仍被阻止，全部来源哈希保持 | `artifacts/compatibility/xc0102-row-20261030-130754-cfb592e2/` |

Core 初次 265 项中，260 项既有测试和 4 项新测试通过，1 项新测试把内置日期格式 14 的显示误写为 `yyyy-mm-dd`。夹具改为显式格式代码后，仅重跑该新增类 5/5；原失败 TRX 保留，不重跑全量套件。

首次 WPF 专项在“清除不存在的单元格不产生新快照”处失败，证据保留在 `artifacts/compatibility/xc0102-row-wpf-20261030-130556-55cd3580/`。这定位到 EditorController 仍使用空字符串编辑：现改为调用 Core 的 `ClearCellsCommand`，保留 20,000 格、桌面坐标上限和预期版本检查。最终 WPF 同时验证真实网格像素、两格一次清除/一次撤销、录入格式和保存重开；截图为 `row-style-inheritance.png`，已检查。

写入测试分别覆盖普通 XLSQ 保存、值快照 XLSX、可移植公式 XLSX；既检查模型重开，也独立检查 row/@customFormat 与 Open XML schema。只变更应用标志也必须触发重写，不能误复用原包。全套 8 条比较路径对新增夹具均无需有损副本；不变路径按来源 SHA-256 保留。

所有 managed 命令使用 `--artifacts-path .\artifacts\build`，未修改同级 ExcelIO。SDK 10.0.401；仍有 NU1900 漏洞源不可达警告。未跑基准、完整回归或发布脚本；WPF 输入是路由事件/合成选择，不是物理输入或真实 OS IME 验收。

## 3. 边界与下一步

- 先前 QuickExcel 写出的文件如果已含 `s=0 customFormat=1`，本轮会按文件声明解释为有效行覆盖，不猜测它是否由旧写入器的尺寸保存路径引入，也不自动改写来源。
- 存储单元格仍持有解析后的样式；此切片不是“修改整行格式并重算所有既有单格继承关系”的编辑模型，没有新增整行样式管理界面。
- 未新增命名样式编辑、主题字体、旋转/缩进/缩小填充或通用 XLSX 保存。已有不支持特性的损失门禁保持。
- 真实 Excel 版本/build、重开与视觉对照仍待补；独立 schema 有效不等于完全 Excel 保真。下一切片继续 XC-0102 的 Normal/命名样式与 XF 关联保留，再按明确范围处理主题字体；XC-0103 的元数据白名单仍独立推进。
