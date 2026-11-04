# 数字格式与真实 Excel 对照

2026-10-30，XC-0101 的以下常用格式范围完成；R/D/E/W，计算值不变。对应提交以本文件首次提交为准。E0 的其他特性样本和 E1 其他任务仍独立验收。

## 支持范围

- 0–30 位小数，`0/#/?` 的必需、可选与占位数字，千分位和尾随逗号缩放，百分比、货币符号和科学/工程计数。
- 正/负/零/文本四分段，比较条件、引号和反斜线字面量、空分段、`@` 文本替换、八种命名格式颜色。颜色覆盖字体颜色但不修改样式；条件分段的负号按 Excel 结果处理。
- 常用会计格式的括号、零横线、数字空位与 `*` 填充。Core 返回宽度无关文本和填充位置，WPF 按可用宽度重复填充字符；复制与自动尺寸用同一内容，自动尺寸不会将已有列宽反复放大。
- 保持既有日期/时间与 1900/1904 子集。公式栏保留可编辑的原始数值或公式，显示和复制显示值使用统一格式器；不将格式文本写回为值。
- 数字分隔符使用当前文化，LCID 可用于日期名称；货币符号以格式声明为准。数字格式代码通过已有样式读写保存，不新增文件版本或必需能力。

边界：`DBNum`、索引 `ColorN`、其他日历仍显示原值；格式代码最多 1024 字符、最多四段。宽度不足维持网格现有裁剪策略；会计 `_` 在文本中预留一个空格，未实现按任意字符精确等宽排版。支持连续数字占位模式，不将电话号码等交错数字/字面量代码视为已验证。内置货币 5–8/42/44 使用美元约定，地区化工作簿应以显式格式代码为准。分数的新增范围见下文，历史对照结果保持原样。

## 2026-11-04 Avalonia 数字格式入口

“开始 → 数字”提供常规、数字、货币、会计专用、短日期、长日期、时间、百分比、分数、科学记数、文本共 11 个带当前单元格示例的快捷格式。会计符号、百分比、千位分隔及增减小数位可直接操作选区；当前格式名称随选中单元格更新。

“其他数字格式…”及右键“设置单元格格式…”（Ctrl+1）打开分类设置窗口：数值/货币的 0–30 位小数、千位分隔、四种负数样式，人民币/美元/欧元/英镑/无货币符号，日期/时间类型，分数与特殊格式（6 位邮编、11 位号码），以及可编辑的自定义代码。单元格格式窗口同时保留“对齐与字体”页。

- 示例使用活动单元格的值或公式结果；空白单元格使用明确标注的 1234.56。预览和取消均不修改文档，确认经 `SetCellStyleCommand` 按当前可见选区一次提交；保留版本、会话、只读、保护及组合输入检查。
- 自定义列表汇集常用代码与当前工作簿已有代码。`G/通用格式`、`[红色]` 等本地化写法保存为可移植的 `General`、`[Red]`；无效或未支持的输入保留在窗口内提示，不静默回退。未修改的导入代码保持原样。
- `@`/内置 49 影响后续用户输入（含数字、布尔、以 `=` 开头的文本）；应用格式本身不转换已有数值或公式，带类型的文件和剪贴板仍保持其类型。
- 新增内置 12/13 与 1–3 位可变分母、1–9999 固定分母，支持整数进位、负数分段、格式色、分子/分母空位。可变分母使用有界连分数寻找最近值；不宣称在所有浮点边界上复刻 Excel 的历史近似选择。超出范围仍回退原值，格式窗口拒绝此类新代码。小数位按钮继续拒绝分数；`TEXT` 函数的独立支持范围不变。
- XLSX/XLSQ 使用既有样式持久化，无新文件版本。所有专项通过后仍须执行发布脚本与生产自检并核对活动包。

针对性验证：`NumberFormatOptionsTests` 覆盖分数、代码校验、本地化、文本输入、预设样式保存重开及撤销；`NumberDisplayTests`、`DateDisplayTests`、`FormatEfficiencyTests`、`TextFunctionTests` 保留相关回归。`scripts/test-avalonia-workspace.ps1 -Mode NumberFormat` 检查真实 Avalonia 菜单/窗口事件、设置恢复、示例、取消、确认与历史，并输出截图；不替代物理 IME/DPI 验收。

规则参考 [Microsoft 自定义数字格式指南](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format)。边界期望由下述 Excel 实测固定。

## Excel 对照与证据

本机 Excel 16.0，build 20527（EXCEL.EXE 16.0.20527.20088），64 位，界面语言 2052，Windows 1900 日期系统。生成器只操作新建夹具，不读取或修改用户工作簿。43 个原始值/格式/显示/颜色期望固定在 `QuickExcel.Core.Test/Fixtures/excel-number-formats.json`，含版本、计算模式与来源 SHA-256。XLSX 原件与截图留在 artifacts。

| 检查 | 结果 | 证据 |
| --- | --- | --- |
| Excel 生成格式夹具 | 43 个固定期望 | `artifacts/compatibility/xc0101-excel-07/` |
| Core 数字与既有日期测试 | 114/114 | `artifacts/compatibility/xc0101-tests/ASUS_LOVESPITE-PC_2026-10-30_13_32_32_net10.0.trx` |
| Excel 来源导入、编辑 B1、写回重开 | 86/86 显示比较，输出独立 schema 0 错误 | `artifacts/compatibility/xc0101-roundtrip-01/display-report.json` |
| 输出再由 Excel 正常只读打开 | 43/43 原值、显示、颜色一致，输入哈希未变 | `artifacts/compatibility/xc0101-reopen-02/excel-reopen.json` |
| WPF 渲染、复制/公式栏/自动尺寸与行样式 | 31/31，截图已检查 | `artifacts/compatibility/xc0101-wpf-05/` |
| Desktop Release 构建 | 通过 | `artifacts/build/bin/QuickExcel.Desktop/release/` |

Excel 会计 `.Text` 的填充空格依列宽和字体改变；比较仅剔除填充位置的重复空格，其余逐字匹配。条件均不命中时的 `#` 个数同样随宽度改变，按溢出标记比较。Excel 用 `Workbooks.Open` 默认 `xlNormalLoad`，未请求修复/提取；[微软说明](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open) 指出对象模型正常打开不会尝试恢复。

真实 Excel 来源仍带 E1 未建模的主题、视图、属性。往返明确使用 `edited.diagnostic-copy.xlsx`，不是普通保存无损声明；保留损失报告。通用 XLSX 保存由 XC-0103 处理。

首次 WPF 检查发现物理鼠标悬停层影响背景取色，改为读取实际格式绘制层；未改动产品悬停行为。失败证据 `xc0101-wpf-01`–`03` 保留。Excel 自动化原型中的本地化颜色和数字/文本赋值已修正，失败目录不算通过。

所有 managed 命令使用 `--artifacts-path artifacts/build`，仅运行相关专项；未跑全量回归、基准或发布。仍有 NU1900 漏洞源不可达警告，无编译错误。未改动同级 ExcelIO。

复现命令（输出目录须新建，旧证据禁止覆盖）：

```powershell
.\scripts\test-excel-number-formats.ps1 -OutputDirectory artifacts/compatibility/new-excel
dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- number-formats artifacts/compatibility/new-excel/excel-results.json artifacts/compatibility/new-roundtrip
.\scripts\test-excel-number-reopen.ps1 -WorkbookPath artifacts/compatibility/new-roundtrip/edited.diagnostic-copy.xlsx -ReferencePath artifacts/compatibility/new-excel/excel-results.json -OutputDirectory artifacts/compatibility/new-reopen
```
