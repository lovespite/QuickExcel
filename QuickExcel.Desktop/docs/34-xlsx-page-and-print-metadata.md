# XC-0105 页面与打印元数据往返

2026-10-30 完成 R/W 子集；对应提交为 `feat: preserve worksheet page and print metadata (XC-0105)`。只读取与保存配置，D/C/E 不适用或未交付；没有打印、分页预览、PDF 或页面设置编辑入口，这些仍属 XC-0401。

## 保存范围

`WorksheetPrinting` 属于 Core 的不可变工作表模型，存储适配在 `XlsxPrinting`。单元格值/样式/视图修改保持页面元数据，写后校验逐项比较；并非依靠不变原包实现的只读保留。

| 配置 | 支持的读写范围 |
| --- | --- |
| 页面设置 | 纸张编号/自定义宽高、方向、百分比缩放、适应页宽高、页序、起始页码、黑白/草稿、批注/错误打印选项、DPI/份数/使用打印机默认值；这些选项的保留不表示已实现相应打印功能 |
| 页边距和选项 | 左右上下及页眉页脚距离（英寸），水平/垂直居中、标题/网格线打印；pageSetUpPr 的 fitToPage/autoPageBreaks |
| 打印区域 | 工作表局部 `_xlnm.Print_Area`，本表绝对单格/矩形/整行/整列及逗号联合，最多 1024 段；解析支持表名内中文、逗号和转义单引号 |
| 重复标题 | 工作表局部 `_xlnm.Print_Titles`，最多一组整行及一组整列；引用使用工作表身份绑定，输出时生成当前名称和 localSheetId |
| 页眉页脚 | 奇数/偶数/首页六段原始格式字符串，differentFirst/differentOddEven/scaleWithDoc/alignWithMargins；保留页码、日期、文件名等控制码，不求值或渲染 |
| 分页 | rowBreaks/colBreaks 的 id/min/max/man/pt；计数验证和写回。不是 QuickExcel 计算出的新分页 |
| 打印机配置 | 仅通过页面设置的内部 printerSettings 关系、对应内容类型、无出站关系且不超过 1 MiB 的不透明字节；重建关系并保持字节，不加载驱动、不访问设备。与普通图片 drawing 关系共存已测试 |

打印引用可表达 Excel 的完整行列上限，仅保存范围元数据，不分配这些单元格，也不提高工作簿内容容量。未知属性、跨表/动态打印公式、未知名称字段、页眉页脚图片及其他未建模部件仍报告并阻止普通有损保存；明确兼容副本才允许移除已报告内容。损坏的页边距、重复定义、分页计数/范围或超预算元数据拒绝读取。

XLSQ 保持 v2；存在打印模型时增加必需能力 `worksheet-printing`，旧读取器明确拒绝未知能力。权威数据仍在内嵌标准 XLSX，不复制第二份 JSON 页面配置。旧 v1/v2 没有页面模型的文件照常打开，已有迁移/未知内容保护保持。

## 编辑边界

- 改值、改样式、改名、插行/插列及一次撤销保留配置；区域、标题和分页坐标随插入更新。改名时打印引用通过当前工作表名重建，删除整张表同时删除其打印名称及关系。
- 没有分页信息时，行列删除使用区域收缩规则，全删重复标题后移除该标题。范围变成不连续或越界时整批拒绝。
- **有分页信息时拒绝删除/移动行列。** 实际 Excel 在本次删除首两行和首两列的夹具中清除了手动分页；早期实现保留移动后的分页，与 Excel 不同。失败证据保留在 `xc0105-excel-delete-01`。当前通过预检拒绝，快照、dirty 和历史不改变，未将该场景标成已对齐。
- 带打印区域/重复标题的工作表拒绝单元格剪切移动；跨表/局部移动对这些名称的影响尚未完整实现。普通改值和格式编辑仍可用。
- 这些边界属于本项 R/W 范围的结构保护，后续打印编辑应补齐具体 Excel 规则后再放宽。

规范参考：[Page setup](https://support.microsoft.com/en-au/excel/page-setup)、[页眉页脚 XML 字段](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.linq.x.headerfooter?view=openxml-3.0.1)、[分页位置与跨度](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.break?view=openxml-3.0.1)、[Excel 分页实现差异](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/77730f13-7af9-42e0-96c7-0b349853617e)。

## 验证

Excel 16.0 build 20527、64 位、语言 2052、自动计算 -4105。只操作新建 artifacts 夹具，普通加载 `xlNormalLoad`，未请求修复。三张表覆盖横向 75%、纵向适应 1×2 页及默认页面；每表比较 48 个页面属性/分页列表，另比较表名。

| 检查 | 结果 / `artifacts/compatibility` 下证据 |
| --- | --- |
| 页面、既有存储/保存/视图/结构及产品规则 | 140/140，`xc0105-related-01` |
| 结构保护加入后的相关专项 | 16/16，`xc0105-structure-final`；最终页面专项 8/8，`xc0105-printing-verified`，含图片关系共存、完整列范围稀疏保留与移动拒绝 |
| 独立比较工具 | 36/36，`xc0105-probe-final`；已支持的 `print-template` 全部八条路径普通保存且独立 schema 无问题；新增 `protected-print-template` 验证未知工作表保护仍报告/阻止 |
| 真实 Excel 来源 | `xc0105-excel-02/excel-printing.json` 与三个 XLSX；来源自身保存重开相等，包含插入及删除参考状态 |
| QuickExcel 改值/样式保存、改名/插行列保存 | `xc0105-roundtrip-final/printing-report.json`：打印模型一致、来源哈希不变、两份输出独立 schema 均 0 问题；分页删除预检保护通过 |
| Excel 读取输出、保存并重开 | `xc0105-excel-reopen-final/excel-printing.json`：48×3 属性/列表与来源相等、重开相等；`xc0105-excel-insert-final`：改名和插入后的区域、标题、分页等与 Excel 自身操作相等 |

页边距的 COM 点数比较取小数点后 8 位，源 XML 英寸数在模型中保持原始 double；其余属性精确比较。最初脚本在关闭所有工作簿后读取 Calculation 返回错误值，已改为在簿打开时记录；`xc0105-excel-01` 不用于计算模式证据。旧夹具断言和 ImmutableArray 相等断言的失败记录保留，不列作通过。

复现：`scripts/test-excel-printing.ps1 -OutputDirectory <新 artifacts 目录>` 创建来源；`dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- printing <来源.xlsx> <新目录>` 执行存储及结构保护；再次运行脚本并传 `-WorkbookPath <输出.xlsx> -ReferencePath <来源 excel-printing.json>`，插入参考加 `-ReferenceState AfterInsert`。

真实 Excel 文件仍含修订/兼容扩展、拼音、含额外名称分组的 app 属性等诊断，首次修改使用显式兼容副本；不宣称整个 Excel 模板任意小改无损。最终 Desktop Release 构建通过，仅既有 NU1900 漏洞源不可达警告。本项没有 UI 修改，未重复 WPF/全量回归、实际打印、性能或发布包装。XC-0106 接续处理不支持公式的受控读取与保留。
