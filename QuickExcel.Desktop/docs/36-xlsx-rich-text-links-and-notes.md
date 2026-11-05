# XC-0107：富文本、超链接与传统批注

2026-11-05 更新：E7 已增加独立富文本编辑器、惰性链接编辑和线程评论桌面入口，当前支持范围与验收见 [83 号记录](83-e6-e7-implementation-log.md)。下文保留 XC-0107 原始验收边界与历史证据；其中“尚不编辑单独 run/评论”的描述已由 E7 声明子集替代。

2026-10-30 完成，提交为本记录所在的 `feat: preserve and display rich text links and notes (XC-0107)`。本项覆盖常见内容 R/D/W、明确的富文本编辑降级及整格字体命令；不是完整富文本编辑器、链接启动器或批注编辑器。

## 范围与契约

| 内容 | 读取和保存 | 桌面显示与编辑 |
| --- | --- | --- |
| 内联/共享富文本 | 保留每个 run 的文本、字体、字号、粗斜体、颜色/主题/索引/tint、字体方案、下划线/删除线及基础字符属性；不把拼音文字混入值 | 绘制常见 run 字体/字号/颜色/粗斜体、普通下划线/删除线；复制仍输出显示文本，自动尺寸使用相同 run 度量 |
| 单格超链接 | 保留外部目标、内部 location、显示文字和 tooltip，写回独立关系；网页、内部地址及 mailto 有真实 Excel 对照 | 保持单元格文字和样式，悬停显示目标/说明。所有地址仅作文本展示，不连接网络、启动浏览器或执行 URI |
| 传统批注（Excel Notes） | 作者、完整文本/run、单元格位置；无外部依赖且仅含 Note 对象的 VML 按字节保留，含几何和可见性 | 单元格红色角标和悬停文本；长内容展示最多 4096 字符并明确截断，完整内容仍保存。源“始终显示”的浮动框暂不复制到桌面，但该标志写回保留 |

富文本内容进入 `CellContent`，注释/链接进入工作表不可变模型。样式、尺寸、计算缓存、XLSQ 身份恢复等复制路径携带新字段。值快照导出保留已有文本 run 和附件，公式仍按原有导出规则处理。

直接输入或粘贴覆盖富文本会整批拒绝；公式栏和网格不进入普通文本编辑。右键 **“转换当前格为普通文本（移除分段格式）”** 明确移除 run 格式，保留原值和单元格样式，一次撤销可完整恢复。清除内容是明确删除文字，普通值编辑/清除不会顺带删掉链接和批注。整格字体命令会将所选字体属性应用到全部 run；填充、边框和对齐保留 run；清除格式移除 run 的字体覆盖。

有链接/批注时暂拒绝结构变换和排序，避免内部地址或 VML 坐标失效；合并不能遮盖非主格的链接/批注。普通值、格式、尺寸编辑可用。此处是当前 R/W 的保护边界，后续复制/结构命令必须实现引用/坐标规则后才能放宽，不能直接删掉这些检查。

XLSQ 仍为 v2，存在 run/链接/批注时增加 `cell-text-annotations` 必需能力；旧读取器明确拒绝。标准内嵌 XLSX 为权威数据。每格文本/分段不超过 32767，单表链接/批注各不超过 65530，单个保留 VML 不超过 1 MiB，并沿用整个文件的解析/取消/内存预算。

## 保留但不完整显示、尚未支持

- 上下标、会计下划线、第三方文件不完整 rPr 的字体回退、混合字号缩小填充和改变文本的数字格式不保证完全对齐；兼容报告定位显示限制，普通保存仍保留已建模数据。
- 拼音、未知 run 属性、修订 UID、复杂/带依赖的 VML、多格 hyperlink ref 等仍报告，普通有损保存继续阻止。新增 `phonetic-rich-text` 夹具验证减少富文本警告后未削弱未知内容保护。
- 不编辑单独 run、不编辑批注、不处理线程评论。线程评论属于不同部件，继续按未知内容保留/报告，不能当成传统批注已支持。
- Excel 批注常见的字体索引 81 保留原索引；桌面使用系统 InfoText 色。对照文件和 [微软 SpreadsheetML 示例](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf)均包含此值；不将任意大于 65 的索引猜测为黑色。

## 验收证据

Excel 16.0 build 20527、64 位、语言 2052、自动计算 -4105。只使用自建 artifacts 工作簿，独立隐藏 COM、宏禁用、外部链接不更新，普通打开、不请求修复；保存并关闭重开。

| 检查 | 结果 / `artifacts/compatibility` 下证据 |
| --- | --- |
| typed 存储、样式/继承、会话、结构、视图、打印及公式保护 | 173/173，`xc0107-related-01` |
| 最终内容/格式专项 | 16/16，`xc0107-text-verified`（内容 10 项、既有格式计算 6 项）；含共享/内联、emoji/换行、RGB tint 不重复应用、显式转换/撤销、继承字体、VML/未知对象保护；最终数量/重复定义/必需属性校验后 10/10，`xc0107-text-guards` |
| 独立工具 | 37/37，`xc0107-probe-03`；后增常见内容普通保存检查及最终相关 7/7，`xc0107-probe-final`，整格字体后复验 1/1，`xc0107-probe-font-verified` |
| WPF | 97/97，`xc0107-wpf-verified`（本项新增 13）；红蓝 run 实际像素、混合字号尺寸、角标/悬停信息路径、显式菜单转换/撤销、普通格恢复编辑和保存重开；`cell-text.png` 已人工查看 |
| Excel 来源 | `xc0107-excel-source-final/excel-cell-text.json`：两表、A1 的 17 字符逐字符 Font/Size/Color/Bold/Italic/Underline、3 种链接和 2 条批注；5 个被验收单元格的记录 |
| QuickExcel 编辑输出 | `xc0107-roundtrip-final-02/cell-text-report.json`：改普通值、链接显示文字、批注所在格的值及整格加粗；模型一致、1921 字节批注 VML 完整保留、来源哈希不变、独立 schema 0 问题 |
| Excel 输出复验 | `xc0107-excel-reopen-final-02/excel-cell-text.json`：上述文字、逐字符格式、地址、说明、作者/批注文本、可见性与 Excel 自身相同编辑参考相符，保存重开相符 |

文本、字符格式、链接及批注属性精确比较。Excel 自身保存 VML 会发生约 0.6 pt 的形状几何取整，因此 Width/Height/Left/Top 单项容差为 0.75 pt（1 个 96-DPI 像素）；原始数值和容差均保留在 JSON，其他字段没有放宽。该验收不是跨 DPI 像素对比或物理鼠标/OS IME 验证。

真实 Excel 的首次编辑仍因修订/拼音等其他未建模内容使用显式兼容副本；不承诺任意源文件无损小改。初次源 VML 精确比较、索引 81、旧富文本拒绝断言、字体 XML 子元素顺序及首次整格加粗继承问题的失败记录均保留，不计入通过。加粗失败说明：创建只有 b 属性的 rPr 会让原本继承单元格字体的文字回退为默认字体；已先物化原字体/字号/颜色再应用命令，并通过新的 Excel 逐字符对照。

复现：`scripts/test-excel-cell-text.ps1 -OutputDirectory <新 artifacts 目录>` 创建来源及 Excel 编辑参考；`dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- cell-text <来源 excel-saved.xlsx> <新目录>` 生成兼容副本；再次运行脚本传入 `-WorkbookPath <输出.xlsx> -ReferencePath <来源 excel-cell-text.json> -ReferenceState AfterEdit`。Desktop Release 构建通过，仅既有 NU1900 漏洞源不可达警告；未运行全量回归或发布新包。
