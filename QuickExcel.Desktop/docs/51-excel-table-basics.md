# XC-0302 Excel Table 基础

2026-10-31 完成声明范围，接续 XC-0303。Table 成为独立文档对象，保留名称、范围、列身份、表头/汇总行、命名样式和自己的筛选/排序状态。XLSX 部件、关系及 XLSQ `worksheet-tables` requiredFeature 一起往返。未关闭 M4/M6 或发布门禁。

## 支持范围

读取和保存有/无表头、有/无筛选按钮、汇总标签、计算列/汇总公式定义、列差异格式。后两类公式仍沿用 XC-0106 整簿只读和原包路径，尚不计算结构化引用。未知表属性、连接、XML 映射或未支持的样式明确报告；未修改可保留原包，修改后的普通保存不能静默丢失这些内容。含公式的值快照会清除公式定义，保留明确导出的字面结果。

数据页提供创建表、调整表范围。创建要求普通文本且不重名的表头；本项边界调整只改变底边，至少保留一行数据，不移动原有内容。普通数据区编辑、完整记录排序、独立表筛选和这些操作的撤销均保留表身份。缩放范围清除此前排序标记，避免留下不再适用的已排序范围。清除筛选保留表和下拉按钮；同一工作表多个 Table 的筛选共同控制可见行。

本项受限编辑保护表头与有定义的汇总格；含汇总行/计算列的范围调整、带表的工作表复制及行列插删/搬移暂拒绝，下一项实现表/列引用联动后再开放。表范围不得与其他 Table、合并单元格或工作表级筛选相交；全簿至多 256 个表，名称/编号全局唯一，表名不能与定义名称冲突。同名表样式须具有同一定义，元数据及条件文本计入历史/保存预算；全表联合筛选扫描不超过 2,000,000 格。

## 样式和显示

60 个内置 Table 样式（Light 1–21、Medium 1–28、Dark 1–11）以命名对象保存。显示时投影主题颜色、粗体、填充和边框，包含行/列条纹、首/尾列、表头/汇总行及角格；不将条纹展开成大量单元格样式。共享边界按区域优先级与线宽处理，筛选/手动隐藏后可见条纹重新计数。较矮的原生 Excel 表头也显示筛选按钮。

支持的自定义表样式为同样的区域、粗体/颜色、纯色填充和基础边框；列差异样式另支持数字格式，显示、筛选候选与条件匹配共用其有效格式。其他差异字体/填充/边框/区域数字格式先保留并报告。直接格式中相对 Normal 改变的属性覆盖表样式；目前模型尚不能完整表达“显式设置为与 Normal 相同”的逐属性覆盖，不能声称任意直接格式都与 Excel 显示一致。原始单元格样式仍保留。

内置定义由本机 Excel `TableStyle.Duplicate` 导出；保存的 tint 是 16 位比例表示，内置显示恢复原百分比，自定义样式使用原精度。表主题颜色按 Excel 实际整数 HLS/灰度行为投影；没有来源主题的新表使用此次导出的默认 Office 主题。已有工作簿使用其原主题，不替换用户主题。此转换只影响 Table，不修改先前普通单元格主题规则。

`BuiltinTableStyles.xml` 和 `DefaultTableTheme.xml` 的来源为 `xc0302-style-catalog-01/catalog.xlsx`，SHA-256 `13AADD91FD2B541156301AD158DE01840E9639817C8960C344DD627F1E2EDABB`。可用 `scripts/export-excel-table-style-catalog.ps1 -OutputDirectory <新 artifacts 目录>` 重建；`xc0302-style-catalog-02` 已验证 XML 语义完全相同。原生样本的数值 150 被 Excel 转为汇总标签文本，模型保留文本类型，不能擅自改为数值。

## 验证证据

Excel 16.0 build 20527 x64、语言 2052；使用任务自有隐藏 COM 实例，记录 PID/版本/文件哈希，禁用宏和外链更新，普通打开和保存重开，不调用修复模式。证据在 `artifacts/compatibility`，原生 XML 和固定期望在测试 Fixtures；生成二进制不入 Git。

| 证据 | 结果 |
| --- | --- |
| `xc0302-excel-01`、`xc0302-probe-04`、`xc0302-output-01` | 8 类表（浅/中/深/条纹/无表头/无筛选/筛选/汇总），3096 项导入和 XLSX/XLSQ 往返，0 差异；Excel 再保存重开与初始期望相同 |
| `xc0302-all-styles-01`、`xc0302-all-style-probe-03`、`xc0302-all-style-output-01` | 60 种内置样式、26820 项颜色/粗体/显示值等检查，0 差异，Excel 输出复验通过 |
| `xc0302-borders-excel-01`、`xc0302-border-probe-03`、`xc0302-border-output-01` | 扩充四边边框与列名，60336 项检查，0 差异；Open XML schema 0 问题，Excel 输出复验通过 |
| `xc0302-excel-edits-01`、`xc0302-edit-probe-01`、`xc0302-edited-output-01` | 独立 Excel 数据编辑和扩表预期，8416 项编辑/撤销/重做/两格式往返，0 差异；输出经 Excel 普通打开、保存重开一致 |
| `xc0302-wpf-06`、`xc0302-ui-output-01` | 15 项真实 WPF 处理器检查通过，包含创建/筛选/排序/扩表/撤销/保存重开和原生较矮表头；桌面创建的文件通过 Excel 普通打开、保存重开 |
| 相关测试、`xc0302-data-operations-01`、`xc0302-workspace-01` | 最终 54 项 Table/持久化/样式/统计/工作区相关测试、85 项数据操作和 153 项工作区 WPF 回归通过 |

Excel 原生样本含本项之外的元数据；比较工具遇到这类保护时使用显式诊断副本并记录原因。自行创建的受控表及列差异格式另有普通 XLSX/XLSQ 保存测试，桌面创建文件也按普通保存验收。不会将诊断副本称为任意来源文件的无损普通保存。

早期失败证据保留：catalog 默认命名空间导致内置样式识别失败，后改为明确命名空间；tint 取整、隐藏行条纹、相邻边框优先级已由独立 Excel 结果纠正。WPF 01/02 暴露较矮表头按钮消失；03/04 的汇总值检查误认原生文本标签为数字，现按原生类型核对。

复现入口：`scripts/test-excel-table-basics.ps1` 的 `-AllStyles`、`-InspectBorders`、`-WorkbookPath`、`-EditBasics`；CompatibilityProbe 的 `table-basics <oracle-json> <新目录>` 与 `table-edits <initial-json> <edited-json> <新目录>`；桌面 `--table-basics-smoke <新目录> [原生文件]`。managed 命令均加 `--artifacts-path artifacts/build`。Release 构建通过，仅已有 NU1900 漏洞源网络警告；未跑全量回归或发布包装。WPF 检查不代表物理输入、OS IME 或完整性能门禁。

参考：[Excel ListObject](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject)、[TableStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.tablestyle)、[Excel 表格式](https://support.microsoft.com/en-us/excel/format-an-excel-table)、[Open XML Table](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.table)。
