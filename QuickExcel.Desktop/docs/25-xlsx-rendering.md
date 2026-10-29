# XLSX 渲染修复（2026-10-29）

使用用户提供的 `笼卡demo.xlsx` 复现。附件未纳入 Git，也未修改；回归夹具在代码中构造。全部修改位于 QuickExcel 仓库，未修改同级 ExcelIO。

## 原因与修改

- XLSX 行高使用 points，Core/桌面尺寸使用 DIP。此前导入直接使用 XML 数字，使 22.8 pt 行只有 22.8 DIP。现在导入除以 0.75（得到 30.4 DIP），导出乘以 0.75；字号继续按 points 处理。仅写出后校验允许行高单位换算的浮点尾差，原始内容保留和修改检测仍严格比较。
- `FormattedText.MaxTextHeight` 小于字体行框时会抑制整个首行，18 pt 中英文混排可出现整格无文字。现在保留自然行框，由单元格裁剪，不再用高度上限隐藏文字；粗体使用 Bold。
- 网格把导入列宽限制为 400 DIP，导致附件 I 列严重变窄。显示和拖动尺寸上限改为 10,000 DIP，保留当前字符宽度换算。超出视口的单元格滚动定位到起始边，避免滚到下一列。换绑较小工作表前先重置编辑位置，避免旧位置在尺寸通知时越界。
- 新增主题颜色、索引调色板、自动色和 HLS tint 解析，用于字体、实色填充和边框。SpreadsheetML 主题顺序按 lt1/dk1/lt2/dk2 映射，系统主题色采用文件中的 lastClr，自定义 indexedColors 覆盖默认调色板。无效引用/tint 拒绝；缺失主题颜色保留报告。导出受支持副本使用具体 ARGB。
- 打开时显示简短兼容性说明，完整技术清单仍可从文件菜单查看。主题颜色已解析，但主题字体/效果/关联仍报告保真限制；没有通过删除全部警告来绕过保存保护。

单位与颜色依据：[Microsoft 行列尺寸说明](https://support.microsoft.com/en-us/excel/change-the-column-width-and-row-height)、[Open XML Color](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.color?view=openxml-3.0.1)。

## 仍有的边界

命名/继承样式关系、富文本分段格式、主题字体和效果、打印设置、文档属性、冻结窗格/选择/缩放等导入视图信息尚未完整建模。附件原始缩放是 55% / 70%，当前采用 QuickExcel 工作区缩放，不承诺与 Excel 像素一致。列宽继续按既有 7 DIP 字符基准换算，并非按每个默认字体的最大数字字宽计算。

未修改工作簿保存为 XLSQ 时完整保留原 XLSX 字节；修改后保存或导出可能损失的未支持内容仍需明确另存受支持副本。旧版本曾把 DIP 数字直接写入 XLSX 行高；本次按文件中 points 的标准含义读取，旧版生成文件的显示高度可能随之变化。

## 验证

只运行相关测试，没有运行全量回归或发布脚本。

- Desktop Release 构建通过，0 错误；存在 NU1900（NuGet 漏洞数据源不可达）警告。
- `TypedStorageTests` 82 项通过，含主题/索引/自定义调色板/tint、无效颜色、尺寸单位往返、浮点尾差和原包保留。结果：`artifacts/validation/xlsx-final/xlsx-storage.trx`。
- `OversizedGridCellTests|M5DimensionTests|DataOperationsTests` 28 项通过。
- WPF `DataOperations` 81 项通过：`artifacts/validation/m5-grid/20261029-083351-440e1188/DataOperations/`。
- WPF `Interaction` 92 项通过（含单元格布局）：`artifacts/validation/m5-grid/20261029-083358-bc151027/Interaction/`。
- WPF `Workspace` 108 项通过（含拖动尺寸、撤销、保存重开）：`artifacts/validation/m5-grid/20261029-083548-87ceec6d/Workspace/`。
- 附件专项 14 项通过：`artifacts/validation/xlsx-rendering-20261029-083721/`。覆盖两张工作表、宽列选择与几何、18 pt 文字像素、拖动/撤销、兼容提示、完整原包保留、受支持副本重开与原件未改动；保留两表及 I 列截图。
- 文字修复前的像素复现位于 `artifacts/validation/xlsx-rendering-before-text-fix/`：22.8 pt 行中的 18 pt 彩色文字像素为 0；修复后为 31 个像素行（本机 DPI），用于证明文字不再整体消失，不作为跨设备字体度量承诺。

专项入口为 `QuickExcel.exe --xlsx-rendering-smoke <仓库 artifacts 下的证据目录> [源xlsx路径]`。省略源路径时只运行合成夹具；传入源文件时读取原件并在证据目录生成快照及测试副本。测试设置与最近文件均隔离在证据目录。WPF 像素和处理路径验证不代表物理鼠标或真实 OS IME 验证。

## 首次打开、滚动和切换工作表的显示一致性

进一步检查发现两处布局生命周期问题，工作簿中的样式数据本身没有被修改：

- 从开始页打开时，`FastGridSurface.Bind` 先设置导入的行列尺寸，之后 `imageGridResized` 首次创建绘图缓冲区又调用文本自动尺寸计算，清除了导入列宽和显式行高。切换工作表重新绑定后恢复文件尺寸，导致首次画面和后续画面不同。现在网格记录宿主已提供尺寸，首次创建/重新创建视口时保留它们；替换 Model 时重置此标记。
- `NotifyRefresh` 在新工作表尺寸应用之前计算表头高度，导致切换后仍使用上一张表的默认行高。现在 `SetModelDimensions` 使用当前尺寸重新计算表头和滚动条边距；视口尺寸变化也通知格式图层重绘。

新增 `MainWindow.SheetSwitchSmoke.cs`。`GridIntegration` 用代码生成 XLSX 并实际打开，覆盖首次布局前绑定、显式行高/列宽、三表往返切换、不同默认行高、背景/字体/边框/合并显示，以及滚动缓存与强制重绘的像素一致性。XLSX 专项也检查原复现副本的两张表往返滚动/切换。没有引入每次滚动清空全部缓存的处理。

修复前的失败证据分别位于 `artifacts/validation/m5-grid/20261029-084709-2c27491e/GridIntegration/`（不同默认行高切换）和 `artifacts/validation/sheet-style-repro/before-initial-layout-fix/`（首次视口覆盖导入尺寸）。修复后：

- Desktop Release 构建通过，输出隔离在 `artifacts/build-sheet-switch`，避免覆盖用户正在运行的程序。构建日志为 `artifacts/sheet-switch-build.log`；包还原关闭 NuGet 在线漏洞审计，本次编译 0 警告、0 错误。
- `GridIntegration` 通过，新增 17 项显示一致性断言：`artifacts/validation/m5-grid/20261029-085817-bc90b128/GridIntegration/`。
- XLSX 专项 28 项通过：`artifacts/validation/sheet-style-repro/final/`。源文件从既有验证包的 `workbook.xlsx` 原样解包到 artifacts，未改动原附件。
- `Interaction` 92 项通过：`artifacts/validation/m5-grid/20261029-085912-b5f19b76/Interaction/`；`Workspace` 108 项通过：`artifacts/validation/m5-grid/20261029-085845-2763c481/Workspace/`。

仅执行相关 WPF 验证，没有运行完整回归或发布脚本，也没有修改 ExcelIO 源码。
