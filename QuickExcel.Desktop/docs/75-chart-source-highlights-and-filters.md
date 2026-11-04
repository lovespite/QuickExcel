# 图表数据源标记与持久筛选

2026-11-04。只修改 QuickExcel；沿用现有对象保护、文档命令、原包保存保护及共享 Skia 渲染路径。

## 使用方式

选中图表后，当前工作表中完整的分类范围显示紫色边框与浅色底，数值范围显示蓝色，系列名称单元格显示红色。共享分类范围去重；标记在图表后方绘制，并使用网格的滚动、冻结、隐藏行列和缩放投影。取消选择或切换工作表时移除标记。选择本身不修改文档。

图表右侧漏斗按钮和“插入 → 对象 → 图表筛选”打开同一个面板。面板包含系列、类别、全选、类别搜索、全部显示、应用和选择数据入口。搜索后的全选只影响匹配的类别；同名类别按源序号分别管理，工具提示显示序号。不同系列具有不同类别排除条件时显示混合状态，未修改的混合项保持原条件。

应用前所有修改只属于面板草稿。Esc 或点击外部取消；应用通过 `SetChartCommand` / `DocumentSession` 提交一条可撤销的记录。过滤只影响图表，不删除单元格或隐藏工作表行列。文档、工作表、选择或版本变化会关闭过期面板；对象保护和只读文档可以查看条件，但不能应用。

## 模型与文件

`ChartSeries.IsFiltered` 保存整个系列的排除状态，`HiddenPoints` 保存完整源范围中从零开始的点序号。数据源引用始终完整保留，重复类别文本不会合并。`ChartData` 保留原系列和点序号，筛选时同步映射点颜色、数值 X 和瀑布总计点。相同分类源的多个系列如果排除条件不同，公共分类轴保留缺口，避免数值移动到相邻类别。全不选显示空数据提示。

普通 XLSX 图表使用 Office 的 `c15:filtered*Series`、`fullRef`、`formulaRef` 和 `categoryFilterExceptions`。可见引用与缓存按筛选结果投影；完整引用和被过滤点的受支持样式仍保留。读取先验证并恢复完整源，然后才读取当前工作簿数据，不信任旧缓存。保存、修改源数据、工作表改名、恢复全选均保持对应关系。重新指定数据范围会重置旧范围的筛选条件。

瀑布/漏斗采用 ChartEx `hidden` 系列标志；类别筛选通过 `urn:quickexcel:chart-filters:1` 的受限扩展保存完整源与样式，并写出筛选后的可见引用。扩展中的完整源、排除序号、可见引用和样式投影必须一致；隐藏内容也经过原有封闭结构验证。XLSQ 增加必需特征 `chart-filters`，旧读者必须拒绝，不能悄悄丢弃条件。

XLSX 保存／导出以 Excel 能正常打开为约束：确认导致 Excel 拒绝文件的状态必须修正或移除。本轮实测表明 ChartEx 筛选扩展可以保留；真正导致原样例打不开的是瀑布系列的显式 `cx:axisId`，以及缺失的图表样式／颜色样式部件。现在写出时移除系列轴引用（保留绘图区坐标轴），补齐缺失样式及必需样式槽；已有样式、主题颜色、完整数据源、类别筛选、点颜色和总计点保留。该规则覆盖普通保存、带版本检查的保存、值快照和可移植公式导出，不依赖用户另外勾选“移除不支持内容”。

旧文件不修改直接另存也会修正这些已知问题，避免原包直存绕过修正。只处理封闭结构验证成功的图表，并按原始图表 XML 和样式部件共同匹配，避免相同 XML、不同配色的图表相互覆盖。其他原包部件保留；已修正文件再次无修改保存保持原包字节不变。修正不改变模型中的筛选结果，不需要把会话标记为有损保存。

协议参考：[Office 完整源引用](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.drawing.chart.fullreference?view=openxml-3.0.1)、[被过滤系列引用](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.drawing.chart.formulareference?view=openxml-3.0.1)、[被过滤点样式](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/9a3eeb15-ec25-4124-86ac-248105d9367e)。未新增生产依赖。

## 验证与边界

仅运行相关测试，没有运行完整回归：

- Release 解决方案构建：0 警告、0 错误。
- `ChartFilteringTests`、`ChartSupportTests`、`DrawingObjectTests`、`E4ReportingTests`：74 项通过。涵盖七种图表的 XLSX/XLSQ 往返、恢复隐藏数据与颜色、原始 Office XML 夹具、横向及带引号的引用、空选择、保护、撤销、不同系列的分类对齐、畸形和未知扩展的原包保护。
- 图表渲染与报表定向测试：23 项通过。包括隐藏首系列后仍使用正确的线条/标记选项，以及所有类型的空筛选渲染。
- Drawing 桌面检查：37 项通过。覆盖真实 Avalonia 路由事件的按钮、面板草稿、应用、搜索、取消、过期状态、保护、撤销重做、文件重开和源范围标记。证据：`artifacts/avalonia/workspace/20261104-005123-3f898755`。
- Grid 检查：90 项通过。证据：`artifacts/avalonia/workspace/20261104-005128-71a980a3`。未将这些自动事件检查当作物理输入或完整平台验收。

额外探针在独立隐藏 Excel 实例中创建、读取和另存本仓库 artifacts 下的样例，未操作用户打开的工作簿。柱形、折线、平滑曲线、饼图的系列/类别筛选被 Excel 正确识别，经 Excel 再保存后也能由 QuickExcel 恢复。证据、输入输出和脚本保留于 `artifacts/chart-filter`。

互操作限制仍需明确：

- Excel 再保存散点图时会删除完整源 `fullRef`，因此该路径不能保证恢复被排除的原始点；缺少完整源且含非连续引用的输入继续明确诊断和原包保护。QuickExcel 自己保存和重开的散点筛选已覆盖。
- ChartEx 类别筛选是 QuickExcel 扩展，保证的是 QuickExcel 的 XLSX/XLSQ 往返，不能宣称其他软件再次保存时会保留其完整恢复信息。之前瀑布／漏斗导出无法被 Excel 打开的问题已修正，见下方验证。Excel 中的类别筛选显示已实测；本机 Excel 忽略 ChartEx 的整个系列 `hidden` 标志，该状态仍可由 QuickExcel 自己恢复。
- Open XML SDK 对 Excel 自产的筛选扩展也报告缺少 `f`、扩展子节点等错误。本次保留原始校验日志和原生 Excel 实测，不宣称筛选样例的 SDK 校验为零错误。
- 外部、动态、多层分类、未知筛选扩展以及未建模的数据点效果仍走原有兼容诊断和原包保护。其他工作簿级兼容诊断也不会因为图表筛选支持而被移除。

## Excel 打开兼容性修正验证

2026-11-04 后续修正，证据位于 `artifacts/chart-export-compat`：

- 隔离的隐藏 Excel 实例正常打开 111 个样例，图表对象及类型均可读取：七种图表 × 五种筛选状态（全选、部分、类别空选、隐藏系列、隐藏系列且类别空选）× 三条写出路径，共 105 个新文件，加上 6 个旧版问题文件不修改直接另存。三条路径同时覆盖绝对、单单元格和双单元格锚点。原始问题文件未覆盖。
- 30 个新 ChartEx 文件通过 Open XML SDK Microsoft365 校验，错误数为零。Excel 导出的 PNG 核验了部分分类筛选后的 20／60／80 数值与空选择；这不代替所有 Office 版本的实测。
- `ChartFilteringTests`、`ChartSupportTests` 本轮最终 50 项通过；同轮 `DrawingObjectTests`、`E4ReportingTests` 32 项通过。新增覆盖旧包修正、其他部件原样保留、重复保存幂等、相同 XML 不同样式、完整／可见系列同步修正及三条写出路径保留筛选。
- Release 解决方案构建：0 警告、0 错误。仅运行相关测试，未运行完整回归；未新增生产依赖或运行时 Excel COM 调用。

## 本地发布

执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，沿用上述定向测试。8 项生产自检通过，默认 Windows 启动指针已更新至 `artifacts/avalonia/releases/20261104-005156-d64f398c/application`，旧发布包保留。发布构建只有既有的 `Program.WebScriptsOnly` 未赋值警告；未提交 Git。

Excel 兼容性修正后再次执行同一发布命令，默认启动指针更新至 `artifacts/avalonia/releases/20261104-012649-dad6e6d3/application`，生产自检通过，旧发布包继续保留。发布构建仍只有上述既有警告。
