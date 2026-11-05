# XC-0305 条件格式与快捷菜单

2026-10-31 完成声明范围。条件格式进入不可变工作表模型、撤销历史、XLSX/XLSQ 和网格显示；XLSQ 声明 worksheet-conditional-formatting requiredFeature。接续 XC-0306，未打发布包，不关闭 M4/M6 门禁。

## 规则与显示

支持单元格八种比较、基础非易变标准公式、重复/唯一值，以及包含、不包含、开头和结尾文本规则。规则保存多区域 sqref、相对公式锚点、优先级、stopIfTrue、文本条件和差异样式。公式条件以第一个区域首格为基准，错误结果不着色；不会执行脚本或把陈旧缓存当作计算结果。已完成重算的结果可复用，否则按当前快照计算需要的标准公式依赖。

高优先级规则覆盖其明确指定的属性，未指定属性可由低优先级规则补充；为真停止后续规则。差异样式支持粗体、斜体、删除线、下划线、字体/纯色填充、四边线和数字格式，显式 false 与未指定区分。共用边框投影到相邻单元格。Excel 的条件差异填充使用 bgColor；普通/table 的 solid fgColor 编码不能直接用于条件格式，实际 Excel 复验曾暴露其白色填充问题，已修正。

重复/唯一值对照覆盖空白与空文本、大小写、布尔值、文本数字、前导零、指数、百分号、千分位、全角数字、Unicode 和通配符。布尔值与文字 TRUE/FALSE 分开。原生证据还确认：条件数字格式改变显示，值筛选仍按基础数字格式匹配，已有筛选引擎保持此语义。

数据页的“条件格式”管理器支持新增、修改、删除、调整优先级和停止规则；整个规则集合一次提交和撤销。规则编辑器提供区域、类型、条件、填充/字体颜色、粗斜体和数字格式；编辑既有规则保留其他已建模差异属性。网格文字、字体、填充、边框、编辑器和自动列宽共用呈现结果。

2026-11-04 增加两色/三色色阶、数据条和 17 种标准图标集。视觉规则具有独立的不可变阈值/颜色/显示模型；高优先级分别决定填充、数据条和图标，普通规则的 stopIfTrue 仍可阻止后续视觉规则。视觉规则本身不允许 stopIfTrue 或差异样式。网格与打印/PDF 共用 Skia 绘制，隐藏值只影响显示，图标为文字留出空间，自动列宽/行高计入图标。

新快照使视口缓存失效；输入、依赖重算、规则修改和撤销后刷新可见单元格。重复值频次按规则缓存。该实现按视口需求计算，不宣称完成全表条件格式的精确脏区域索引或大型性能门禁。每表上限 1,024 条规则、4,096 个区域、4 MiB 规则元数据，公式 8,192 字符；一次投影计算预算 2,000,000 步，当前快照重复值区域累计预算 2,000,000 格。超限或不支持的依赖显示状态栏/日志与悬停诊断。

## 结构变换与保护边界

行列插删调整作用范围和公式；删除原锚点时先取首个存续单元格的相对公式，避免错误地产生 #REF!。整段规则删除后移除规则。工作表复制、重命名、删除，名称与表改名沿用 AST 引用变换，复制后的规则集合独立。按值排序支持每个相关规则覆盖全部待排序行的情况。

以下范围明确受控：

- 原生 top10/aboveAverage/timePeriod、自定义图标组合、x14 新图标集及未建模扩展条件规则尚未支持。快捷排名/平均值使用下述普通公式规则，不扩展原生类型的导入支持。未知高优先级规则可能抑制其他规则，因此整张表的条件集合随原包保留并报告，不能只丢弃未知条目后继续显示低优先级规则。
- 含未知条件规则的工作簿不允许通过管理器覆盖集合；不变保存保留原包，改变后的存储仍遵循兼容性报告与显式诊断副本路径。
- 行列搬移、部分规则行排序及条件颜色排序尚未支持。数据验证、图片、批注等既有结构保护仍生效；并非所有混合特征工作簿都能插删。
- 条件格式区域的完整/格式粘贴、格式刷、复制式填充、序列填充和清除格式尚未实现规则拆分，明确拒绝整批操作。只贴值/公式保留目的规则；规则通过管理器调整。剪切沿用结构保护。
- 差异字体名称/大小、图案填充等未建模样式进入保留/报告路径；本项不宣称完整 Excel 条件格式。

## 验收证据

### 2026-11-04 Avalonia 快速条件格式菜单

“开始 → 样式 → 条件格式”采用分类子菜单：突出显示单元格规则、最前/最后规则、数据条、色阶、图标集、新建、清除和管理。数据页保留原管理器入口。“发生日期”保留禁用提示，其易变日期语义尚未实现。

快捷创建支持大于、小于、介于、等于、包含文本、重复/唯一值、前后 N 项/百分比、高于/低于平均值。对话框固定创建时的连续矩形选区，提供六种配色/边框样式和本地样例；取消不写入，确认作为一个可撤销命令插入最高优先级。比较值支持数字、文本和以 `=` 开头的公式；完整自定义规则直接打开原规则编辑器。编辑既有规则保留会话 ScriptId。

排名/平均值通过已有非易变标准函数生成 `expression` 规则，阈值按原选区动态计算，使用绝对统计区域和相对当前格，XLSX/XLSQ 保存重开无需新文件类型。排名项数为 1–1000、百分比为 1–100；百分比向上取整，并列阈值一起匹配，仅数字参与比较。无数字或公式计算错误不着色；依赖错误、未支持的计算和资源预算仍服从已有公式呈现契约，不宣称完整原生 top10/aboveAverage 语义。管理器中这些快捷规则显示为“公式”。

“清除所选单元格的规则”支持多个选区，在 Core 中按矩形相减拆分规则范围；原锚点被移除后重定位两条相对公式，保留规则身份、优先级、差异样式和停止标志。“清除整个工作表的规则”只处理活动表。区域/元数据超限、工作表保护、原包未知规则均原子拒绝；普通“清除格式”、剪贴板及脚本 API 的既有保护不放宽。所有入口保留 IME composition、原格草稿提交和修订检查。

视觉快捷菜单提供 12 种数据条（六色 × 渐变/实心）、12 种两色/三色色阶和 17 种标准图标集，点击即应用为一个撤销步骤。各分类的“其他…”与管理器可编辑作用范围和阈值；支持数字、百分比、百分位、非易变公式，色阶另有最小/最大值，数据条另有自动最小/最大值。阈值公式须使用绝对单元格引用（例如 `$A$1`），且遵守原公式绑定/计算边界。高级数据条可修改最小/最大长度、正负颜色、边框、轴位置/颜色、左右方向及是否显示值；图标支持反转、仅图标显示和阈值等号选择。上下文方向采用当前从左向右工作表布局。

数值分布按工作表/规则范围缓存，稀疏枚举有内容的单元格；空白、文本、布尔和错误值不参与分布、不生成视觉标记。百分位使用线性插值。新快照使分布失效，数值或阈值依赖改变后刷新。当前快照累计分布枚举预算为 2,000,000 格，沿用公式计算预算；不把整列区域展开为稠密数组。逆序阈值和超过预算产生诊断，不宣称完成大表性能门禁。

XLSX/XLSQ 读写标准 `colorScale`、`iconSet`，数据条同时写基础兼容规则和关联的 x14 扩展，保留实心/渐变、负值、边框、轴和方向。x14 规则须以 GUID 和同一 sqref 对应基础数据条；未知/重复/孤立扩展进入整集合保留路径。基础图标集为 `3Arrows/3ArrowsGray/3Flags/3TrafficLights1/3TrafficLights2/3Signs/3Symbols/3Symbols2/4Arrows/4ArrowsGray/4RedToBlack/4Rating/4TrafficLights/5Arrows/5ArrowsGray/5Rating/5Quarters`。代码绘制对应方向、形状、颜色和等级，不嵌入 Excel 位图素材。

原生 Excel 复验发现扩展 URI 必须保留 `{78C0D931-6437-407d-A8EE-F0AAD7539E65}` 中的小写 `d`：全部大写虽通过 Open XML schema 校验，Excel 普通打开仍会忽略关联的高级数据条设置。写出采用原生大小写和限定的 `x14:id`，导入容忍 GUID 大小写。独立 XML fixture 固定原生扩展结构及未知属性保留行为。

定向验收（未执行全量回归，不替代物理输入验证）：

| 检查 | 结果与证据 |
| --- | --- |
| Core 条件规则、视觉规则、快捷菜单命令与相关 ScriptP2 回归 | 41 项通过；`artifacts/test-results/conditional-visuals/conditional-visuals.trx`。覆盖排名/并列/百分比/动态依赖、全部预设 XLSX/XLSQ 往返、负值轴/反向/阈值/停止规则、稀疏区域、结构变换、局部拆分/重叠选区/公式锚点、保护/未知内容与身份保留 |
| Skia 视觉像素与打印会话 | 6 项通过；同目录 `conditional-rendering.trx`。覆盖数据条/轴、隐藏值、17 种图标裁剪、打印会话 |
| `scripts/test-avalonia-workspace.ps1 -Mode ConditionalFormat` | 31 项真实处理器检查通过；`artifacts/avalonia/workspace/20261104-141403-8f6064fd/` 包含菜单、配色预览、高级数据条对话框、网格截图和导出文件。覆盖输入校验、取消、撤销重做、保护、合成 composition、全部三类视觉菜单、三档切换五档图标、编辑身份和文件重开 |
| 原生 Excel 普通打开、重算、保存、重开 | `artifacts/compatibility/conditional-visuals-20261104-final/`；源哈希不变、保存前后稳定；`verify-conditional-visuals` 7 项通过，六条规则元数据一致，QuickExcel 源文件 Open XML schema 零问题。包括负值渐变/实心数据条、边框/隐藏值、三色色阶和三种图标集 |

生产自检同时验证三种视觉规则投影、XLSQ 保存重开及包含规则的 PDF 输出。发布使用 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，成功后由 `artifacts/avalonia/active-windows.json` 指向通过自检的新包。脚本 API 的详细视觉配置仍按 [API P2 边界](79-script-api-redesign.md) 处理。

本次交付包：`artifacts/avalonia/releases/20261104-141427-092fa74e/application`，14 项生产自检通过；已确认 active-windows 指针及 release.json SHA-256 匹配，`Start-QuickExcel.cmd` 使用该自包含包，诊断测试入口不进入发行包。

视觉读写规范参考：[数据条扩展 CT_DataBar](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/88c5ca0a-3a34-4fad-8c91-fdb4550d3fe1)、[阈值 CT_Cfvo](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/15094f93-f587-4f4e-be74-1fbdc93da768)、[条件规则 CT_CfRule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/025ea6e4-ad42-43ea-a016-16f4e4688ac8)。

### 历史基础规则验收

本机 Excel 16.0 / build 20527，x64，语言 2052；文件版本 16.0.20527.20088。脚本创建自有隐藏 COM 实例，禁用宏及外部更新；使用普通 Open、CalculateFullRebuild、SaveAs、关闭重开，未使用修复模式。打开前记录文件哈希，结束仅退出自有实例。证据位于仓库 artifacts/compatibility，关键原生包 XML、显示结果、版本及 SHA-256 固化于测试 fixture excel-conditional-formatting.json。

| 证据 | 结果 |
| --- | --- |
| xc0305-excel-04 | 比较、文本、重复/唯一值、重叠优先级/停止、共享边框、数字格式六张表；初始、编辑、插行、删除首行、复制表五阶段保存重开稳定 |
| xc0305-probe-08 | 131,982 项原生导入、命令变换、撤销重做及 XLSX/XLSQ 往返检查，零差异；所有模型输出 XML schema 零问题 |
| xc0305-output-{initial,edited,insert-row,delete-first,copy-sheet}-02 与对应 verdict-02.json | 五阶段 QuickExcel 输出由 Excel 普通打开、重算、保存重开；29,489 项原生属性和显示检查，零差异 |
| xc0305-wpf-06 | 17 项真实 WPF 入口检查：对话框、管理器、优先级/停止、依赖刷新、重复值更新、撤销、自动列宽、字形像素可见性、普通保存重开及原生导入；截图已检查 |
| xc0305-ui-output-03 / xc0305-ui-verdict-03.json | 桌面保存产物在 Excel 重算重开，121 项显示复验零差异，源哈希不变；设置明确列宽以排除 Excel 窄列的 #### 显示 |
| xc0305-workspace-02 | 153 项工作区 WPF 回归通过 |
| 最终相关测试 | 205 项条件格式、验证、表/引用、名称/结构、条件统计、剪贴板、填充和格式效率测试通过；其中条件格式专项 12 项 |

原生文件的本项以外 UI/扩展元数据仍走显式诊断副本；桌面创建的条件格式文件另按普通保存验收。Release 构建成功，仅现有 NU1900 漏洞源网络警告。未运行全量回归或发布脚本。WPF 合成事件不是物理输入或 OS IME 证明。

复现入口：scripts/test-excel-conditional-formatting.ps1（新原生矩阵或 -WorkbookPath 复验）、CompatibilityProbe 的 conditional-formatting / verify-conditional-formatting / verify-conditional-model，以及桌面 --conditional-formatting-smoke。托管命令始终使用 --artifacts-path artifacts/build。早期 dxf 字体元素顺序、填充色、重复值数字/通配符和窄列差异证据保留，不覆盖失败记录。

参考：[Microsoft 条件格式说明](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel)、[OOXML 条件格式扩展](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/025ea6e4-ad42-43ea-a016-16f4e4688ac8)、[FormatCondition 边框接口](https://learn.microsoft.com/en-us/office/vba/api/Excel.FormatCondition.Borders)。

## XC-0606b 发生日期 / A date occurring

1.28.0 源码支持今天、昨天、明天、含今天的最近七天、本/上/下周、本/上/下月。星期日起算；日期用工作簿 1900/1904 序列号，数值不要求日期格式，明确的年优先日期文本及数字文本按现有稳定转换规则匹配，模糊地区日期文本保持未转换。空白、布尔和错误不匹配。主界面“突出显示单元格规则 → 发生日期”及管理器均可编辑，优先级、为真停止、差异样式和对象身份沿用既有规则。

Date periods are a separate model kind. Their canonical Excel timePeriod/formula pairs are validated on import and regenerated on export; unknown periods or modified formulas retain the original package. General conditional formulas still reject arbitrary volatile functions. XLSX/XLSQ use the same representation. Existing script handles retain the date period on metadata changes; getType reports the unsupported script exposure instead of pretending it is a text rule.

样式缓存同时绑定不可变快照和本地日期；桌面一分钟检查日期变化并重绘，不创建编辑历史。打印/PDF 继续使用同一条件投影。模拟跨日测试验证同一快照结果改变；未将其描述为等待到真实午夜的实机证据。标准函数和脚本依赖继续遵守既有计算/新鲜缓存约束，不为条件格式授权脚本。独立 Excel 和定向验证见 83 号记录。
