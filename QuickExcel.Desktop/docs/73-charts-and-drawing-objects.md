# 图表样式、图表类型和对象操作

2026-11-03。此记录更新 [57 号记录](57-e4-reporting-and-capacity.md) 中图表子集及图片旋转边界。修改只在 QuickExcel 仓库；未修改 ExcelIO、脚本隔离或未知内容的原包保护。

## 图表支持

| 类型 | 数据和显示 |
|---|---|
| 簇状柱形 | 多系列；柱间距、系列重叠、系列和数据点颜色 |
| 折线 | 多系列；点标记、线宽、断点、数据标签 |
| 平滑曲线 | 折线的平滑模式；同一系列内连续有效点绘制曲线 |
| 散点 | 数值 X/Y 轴；标记、直线或平滑连线；独立 X/Y 范围和数字格式 |
| 饼图 | 单系列；分类配色和标签 |
| 瀑布图 | ChartEx；增减累计、显式总计/小计点、连接线及对应配色 |
| 漏斗图 | ChartEx；单系列非负值、按源顺序居中分段、分类颜色和标签 |

“插入 → 图表管理”可创建、编辑数据范围、切换类型、设置标题/图例、删除。首行是系列名称，首列是分类，散点图首列是数值 X。饼图、瀑布图、漏斗图需要两列。瀑布图总计点使用从 1 开始的逗号分隔序号；总计点从零绘制，并将后续累计基准设为该值。

数据仍要求本工作簿直接单行/单列单元格引用，每系列最多 4,096 点、每图最多 16 系列、每表最多 128 图。保存按完整源序号重建缓存；显示读取当前快照，保留隐藏行/列、筛选和总计点的源序号关系。公式未重算时不使用旧缓存。散点图的非数值 X/Y 与折线缺失值形成断点。ChartEx 当前仅建模可见数据、单层分类和单系列瀑布/漏斗布局。

## 附加样式

读取并原样保留包内无附加依赖的 `chartStyle` / `chartColorStyle` 部件，以及图表中的受支持 DrawingML。支持主题/显式颜色、透明度和常用明暗变换、循环调色板、标题/坐标轴/图例字体、背景、边框、主网格线、线宽、标记、柱间距和重叠参数。每种样式部件最多一个，每部件最多 512 KiB。

编辑数据或移动、旋转、缩放对象时，原图表格式节点和附加样式字节保持；仅更新已建模字段与缓存。切换图表类型重建该类型的结构，保留通用颜色和文字外观，清理旧类型的专属参数。新瀑布图输出明确的增/减/总计调色板。局部数据点颜色不覆盖其他瀑布点的自动颜色。

`artifacts/examples/sample1-chart1.xlsx` 是普通簇状柱形图。旧适配器拒绝它的两条附加样式关系，其后还会拒绝字体/主题绘图属性及 `gapWidth=219`、`overlap=-27`。这些图表内容现已支持。该工作簿仍有与图表无关的兼容性提示，保持原来的保存保护；验证使用独立导出副本，未修改输入文件。

外部数据/样式、样式的附加依赖、未知扩展、阴影/渐变等未建模绘图效果，及堆积、水平、组合、3D、趋势线、误差线、日期/对数/多组轴、动态/名称引用等仍走明确诊断及原包保留路径。这里不是所有 Excel 图表及主题效果的完整实现。自动刻度、字体布局和平滑插值由共享 Skia 渲染器完成，不宣称与 Excel 逐像素等价。

ChartEx 结构参考 [Open XML SDK 官方 schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_microsoft_com_office_drawing_2014_chartex.json) 及 [Microsoft Office 绘图扩展规范](https://learn.microsoft.com/nl-nl/openspecs/office_standards/ms-odrawxml/2ea3f228-fe39-4f55-b8ec-cee89596e926)。未增加生产 SDK 依赖。

## 图片和图表对象

- 单击选择对象，显示边框、八个尺寸控制点和顶部旋转控制点。拖动内部移动，拖动边框控制点调整尺寸，拖动顶部圆点旋转。
- Shift 移动约束到一个方向；Shift 拖动角点保持比例；Shift 旋转按 15° 吸附。方向键微移 1 像素，Shift + 方向键微移 10 像素。
- Esc 取消当前预览；再次 Esc 退出对象选择。鼠标捕获丢失、切换工作表或文档版本变化同样取消预览。
- Delete 删除所选对象；撤销和重做恢复完整对象。图片/图表管理对话框自动选中当前对象，并可输入精确位置、尺寸、角度及随单元格变化规则。

选择和拖动预览不写入文档。松开鼠标后，通过 `TransformDrawingObjectCommand` / `DocumentSession` 一次提交完整变换，形成一条撤销记录。提交核对文档、工作表、版本和原锚点。对象选中时停止单元格编辑及公式栏写入；保护、只读、输入法组合、未完成的编辑提交继续生效。

坐标使用 96 DPI 像素与 EMU，旋转保存为 DrawingML 的 1/60,000 度。命中测试在对象局部坐标进行；冻结窗格、滚动和缩放使用同一投影。网格、报表/PDF 和打印共享对象旋转，旋转后的外接矩形参与裁剪及页面范围计算。支持 twoCell、oneCell、absolute 三类锚点。裁剪、翻转、分组及单元格内图片仍保持原有边界。

XLSQ 添加 `worksheet-charts-v2` 和 `drawing-transforms` 必需特征，避免旧读者静默丢失新增语义。

## 验证

仅执行修改区域相关测试，未运行完整回归：

- Release 全解决方案构建：0 警告、0 错误。
- `ChartSupportTests`、`DrawingObjectTests`、`E4ReportingTests`：50 项通过。包括独立 ZIP/XML 样式及 ChartEx/散点夹具、源数据编辑、改名、类型切换、XLSX/XLSQ 往返、拒绝未知依赖、保护/撤销及旋转锚点。
- `ChartRenderingTests` 和 `RenderingTests.Report`：5 项通过，验证数值 X 位置、平滑曲线、局部瀑布配色、图像旋转和报表输出。
- `WorksheetPrintingTests`：8 项通过，补查旋转对象范围涉及的分页路径。
- `test-avalonia-workspace.ps1 -Mode Drawing -SkipBuild`：22 项通过。实际 Avalonia 路由事件验证选择、125% 缩放下冻结/滚动窗格中的移动、缩放、旋转、撤销/重做、Esc、捕获丢失、版本变化取消、删除、保护及保存重开。证据：`artifacts/avalonia/workspace/20261103-133333-cd3b7c7b`。
- Grid 模式：77 项通过。证据：`artifacts/avalonia/workspace/20261103-132334-02343e16`。
- 独立探针用 Open XML SDK Microsoft365 校验七种输出图表和样例的受支持导出副本，新增结构错误均为 0；PNG 及 XLSX 留在 `artifacts/chart-support/evidence`，样例、平滑曲线、瀑布和漏斗已目视检查。

UI 事件检查不是物理鼠标、多显示器 DPI 或真实 Excel 逐像素兼容性验收；现有物理 IME、可访问性、设备和跨平台门禁未因此关闭。测试探针及 DrawingSmoke 不进入生产包。

## 本地交付

执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，沿用上述定向测试结果，避免发布脚本重复运行全量回归。生产包为 `artifacts/avalonia/releases/20261103-133506-0132d638/application`，已通过 8 项生产自检并更新默认 Windows 启动指针。原发布包仍保留。生产自检覆盖资源完整性、文件保存重开、隔离 JSC/TypeScript worker、Monaco、PDF 和资源清理。发布构建仍有既有的 `Program.WebScriptsOnly` 未赋值警告；解决方案常规 Release 构建无警告。此次工作未提交 Git。

## 横轴标签布局修复（2026-11-04）

柱形、折线、平滑曲线和瀑布图的分类标签改用实际字体排版尺寸。原先固定按 60 像素跳过标签，并将标签居中到多个分类的区间，造成短标签缺失和位置偏右；现在各标签始终锚定自己的数据点，空间充足时显示全部分类。

密集时按水平、45°、90°的顺序尝试完整显示，90°仍会重叠时才按间隔减少标签。旋转后重新计算绘图区边距，为首尾文字、完整文字高度和底部图例预留空间。网格、报表、PDF 和打印继续共用此 Skia 路径。

验证：Release 解决方案构建无警告、无错误；`ChartRenderingTests` 与 `RenderingTests.Report` 共 15 项通过，覆盖四种分类图的对应点居中、三级角度选择、90°后的间隔显示、完整标签内容、长标签、大字号及底部图例边界。未运行完整回归。只读加载原样例，已目视检查修复后的 `artifacts/chart-axis-fix/sample1-fixed.png` 与四种布局的 `artifacts/chart-axis-fix/label-layouts.png`。

执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，8 项生产自检通过，默认启动指针已更新至 `artifacts/avalonia/releases/20261104-002958-93ac10f4/application`。旧发布包保留；发布构建仅有上述既有的 `Program.WebScriptsOnly` 警告。此次修复未提交 Git。
