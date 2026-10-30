# XC-0104 工作表视图与可见性

2026-10-30 完成声明子集；对应提交为 `feat: persist workbook views and sheet visibility (XC-0104)`。冻结、缩放、活动表/格及工作表可见性进入 Core 不可变模型，XLSX/XLSQ 保存后可在 Excel 恢复。此记录更新 18 号记录中冻结仅为会话状态的旧契约；筛选仍为会话状态，等待 XC-0301。

## 支持范围

| 维度 | 范围 |
| --- | --- |
| R / W | 单工作簿窗口、普通工作表视图；冻结行/列或双轴、frozen/frozenSplit、10%–400% 缩放及普通缩放、每个窗格的活动格和单一矩形选区、首个滚动格；活动表及 visible/hidden/veryHidden。窗口位置/大小原单位保留，不移动桌面窗口 |
| D | 可见表标签、每表缩放/活动格/矩形选区和滚动位置；已有网格投影最多冻结 10 行，列限于前 256 列。超限视图保留并在状态栏/兼容报告说明，不声称完整显示 |
| E | 冻结/取消冻结、每表缩放、普通隐藏/取消隐藏；至少一张可见表。隐藏活动表自动选择可见表，按稳定 SheetId 映射标签、地址跳转及切表快捷键 |
| C | 不适用；隐藏表仍在领域和计算模型中，公式可以引用 |

普通“取消隐藏”只列出并接受 hidden，不能通过输入 veryHidden 名称绕过。查找/替换不遍历隐藏工作表；查找的“包含隐藏”选项仍指行列。工作表隐藏不提供加密或访问安全。

非冻结拆分、多窗口、组合选中的工作表、分页/页面布局视图、非连续选区及超出 Core 坐标预算的视图继续诊断。原包不变可保留，编辑后普通保存阻止有损改写，显式兼容副本说明损失。未知视图属性、非默认窗口选项也进入诊断，不悄悄改成默认并宣称无损。

## dirty、历史与格式契约

- 冻结、缩放、可见性使用文档命令，产生 dirty 和一次可撤销历史，不重新计算公式；撤销/重做同步网格和标签。
- 切换工作表、活动格和滚动属于导航，通过会话同步更新快照，不产生内容修订、dirty、历史或重算。显式保存和下一次内容保存捕获这些状态；仅导航后关闭不强迫保存提示。恢复视图本身不修改状态。多选操作只持续记录活动格，原文件矩形选区可读写和恢复。
- 插入冻结边界之前的行/列扩展冻结范围，删除缩小范围；删除一条冻结轴后折叠对应窗格，完整删除冻结区后取消冻结。活动格坐标保留。双轴冻结前插入一行的滚动/活动格行为已与真实 Excel 对照，其他复杂移动场景不据此声称逐项对齐。
- XLSQ 仍为 v2；非默认视图/可见性增加必需能力 `workbook-views`，旧读取器会拒绝未知必需能力。标准内嵌 XLSX 是权威数据，state.view 只镜像活动表/格/缩放。旧 v2 没有标准工作簿视图且未声明新能力时，导入旧 state.view（缩放比例转百分比）；没有显式 bookViews 的合法默认窗口仍可读取。v1 迁移及原文件保护保持。
- 所有克隆、样式/结构/合并命令保持视图/可见性；写后校验比较视图、活动表名称与窗口字段。稳定身份不写入普通 XLSX，打开时重新绑定。

规范参考：[SheetViews](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetviews)、[Pane](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pane)、[Selection](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.selection)、[冻结窗格位置语义](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bb178944-7838-4a58-9c16-e0c63d428a43)。

## 验证与复现

本机 Excel 16.0 build 20527，64 位，语言 2052，自动计算；仅操作新建隔离夹具。使用普通加载 `xlNormalLoad`，未请求修复；不将不存在的 RepairMode 属性作为证据。来源包含普通、双轴冻结、仅冻结行、仅冻结列四种视图，以及 hidden / veryHidden 共六张表。

| 检查 | 结果 / `artifacts/compatibility` 下证据 |
| --- | --- |
| 视图、既有存储、保存、样式/结构回归 | 120/120，`xc0104-related-final` |
| 最终视图与旧元数据专项 | 9/9，`xc0104-view-final`，包括无 dirty 导航/滚动、一次撤销、原包保护、旧 v2 迁移与合法隐式窗口 |
| 查找替换及产品规则 | 13/13，`xc0104-product-rules`，包括隐藏表不被查找替换 |
| 独立比较工具 | 34/34，`xc0104-probe-tests` |
| 实际 WPF 视图/保存处理器 | 74/74，`xc0104-wpf-03`；本项新增 24 项，含 veryHidden 拒绝、远处 G501 活动格和视口保存重开；截图 `workbook-view.png` |
| 工作区与产品相关回归 | 151/151，`xc0104-workspace-02`；41/41，`xc0104-product-03`。冻结夹具改走文档命令；合并底色取样移开选区高亮。旧失败证据保留 |
| Excel 来源保存重开 | `xc0104-excel-01/excel-views.json`：活动表、可见性及四表冻结/缩放/活动格/滚动位置完全一致 |
| QuickExcel 修改并保存 | `xc0104-views-roundtrip-01/views-report.json`：来源哈希不变，视图模型一致；普通改值及插行两份输出独立 schema 校验均 0 问题 |
| Excel 读取 QuickExcel 输出并保存重开 | `xc0104-excel-reopen-01/excel-views.json`：与来源视图逐项相等、再次重开相等；`xc0104-excel-insert-01`：插入第一行后的视图与 Excel 自己插行的固定结果一致 |

复现：`scripts/test-excel-views.ps1 -OutputDirectory <新 artifacts 目录>` 生成 Excel 来源与插行参考；`dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- views <来源.xlsx> <新目录>` 生成普通修改/插行输出；再次执行脚本并指定 `-WorkbookPath <输出.xlsx> -ReferencePath <来源 excel-views.json>`，插行另加 `-ReferenceState AfterRowInsert`。

真实 Excel 来源仍带打印、修订扩展等未建模内容；本次首次修改采用明确的兼容副本并记录损失，不宣称整个模板无损。WPF 验证调用真实处理器、采用注入对话框，非物理鼠标/IME 验收。Release 构建通过，仅既有 NU1900 漏洞源不可达警告；未跑全量回归、性能或发布包装。下一项 XC-0105，E1/E3 尚未整体完成。
