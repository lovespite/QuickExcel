# 数字格式与真实 Excel 对照

2026-10-30，XC-0101 的以下常用格式范围完成；R/D/E/W，计算值不变。对应提交以本文件首次提交为准。E0 的其他特性样本和 E1 其他任务仍独立验收。

## 支持范围

- 0–30 位小数，`0/#/?` 的必需、可选与占位数字，千分位和尾随逗号缩放，百分比、货币符号和科学/工程计数。
- 正/负/零/文本四分段，比较条件、引号和反斜线字面量、空分段、`@` 文本替换、八种命名格式颜色。颜色覆盖字体颜色但不修改样式；条件分段的负号按 Excel 结果处理。
- 常用会计格式的括号、零横线、数字空位与 `*` 填充。Core 返回宽度无关文本和填充位置，WPF 按可用宽度重复填充字符；复制与自动尺寸用同一内容，自动尺寸不会将已有列宽反复放大。
- 保持既有日期/时间与 1900/1904 子集。公式栏保留可编辑的原始数值或公式，显示和复制显示值使用统一格式器；不将格式文本写回为值。
- 数字分隔符使用当前文化，LCID 可用于日期名称；货币符号以格式声明为准。数字格式代码通过已有样式读写保存，不新增文件版本或必需能力。

边界：分数、`DBNum`、索引 `ColorN`、其他日历仍显示原值；格式代码最多 1024 字符、最多四段。宽度不足维持网格现有裁剪策略；会计 `_` 在文本中预留一个空格，未实现按任意字符精确等宽排版。支持连续数字占位模式，不将电话号码等交错数字/字面量代码视为已验证。内置货币 5–8/42/44 使用美元约定，地区化工作簿应以显式格式代码为准。

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
