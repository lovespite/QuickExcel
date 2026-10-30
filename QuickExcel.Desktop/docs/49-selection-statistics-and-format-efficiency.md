# XC-0210 选区统计与格式效率

2026-10-30 完成声明范围。E2 十项均已有对应实现、相关测试及真实 Excel 对照，接续 E3 的 XC-0301。此记录不关闭 M4/M6 或发布门禁。

## 选区统计

状态栏显示选区求和、平均和非空计数。仅遍历不可变快照中的已存单元格，不展开整行/整列空白格；多选区取并集，重叠只计一次。排除手动隐藏行、手动隐藏列和筛选隐藏行。计数包含文本（含空文本）、布尔和错误；求和/平均仅计 Number，不把文本编号、布尔或错误转换为数字。零数字时只显示计数；数值溢出显示破折号并提示，平均保留可计算的有限结果。

公式使用已有缓存，不由选区变化触发重算。旧缓存计入但明确标记“含待重算结果”，缺失缓存不计值并在提示中列数；错误数量单列。筛选结果过期或不能求值时暂停统计。80 ms 防抖后在后台计算，切换选区/文档或关闭窗口取消旧任务，结果须匹配当前快照才发布。最多 64 个选区、2,000,000 次范围比较；达到上限说明原因，不把部分总和显示成完整结果。求和采用补偿求和。

该产品统计口径用真实 Excel `AGGREGATE(9/1/2,7,range)` 和 `SUBTOTAL(103,range)` 对照数值和计数。不是直接读取 Excel 状态栏 UI，也不宣称隐藏列/错误的产品选择与 Excel 状态栏所有设置相同。全可见、手动隐藏、筛选与手动隐藏并存三个状态均有固定结果。

## 格式命令

数字组增加“增加/减少小数位”，逐格调整原格式中的小数占位符，保留百分比、货币、千分位、缩放、科学计数、颜色、条件和文本段。General 根据当前数值的 15 位有效数字显示构造格式；自定义 `#`/`?` 可选占位保持原种类。支持 0–30 位小数，单次最多 20,000 格、总样式最多 4096。日期/时间、分数、文本和不能明确解析的格式整批拒绝；不会部分写入。只改变显示，不舍入值、不改公式或文字类型。修正工程科学计数零值所需的整数占位补零。

常用格式入口包含常规、两位小数、百分比、日期、千分位、人民币货币和科学计数。格式刷先点击采样当前格，再选择矩形目标、点击“应用格式”；Esc 取消。采样期间文档发生编辑会使其失效。复用内部剪贴板的格式粘贴和原子撤销，保留内容/公式/缓存，不接管系统剪贴板；同文档跨表可用。合并、图片、隐藏维度、富文本等沿用 [40](40-internal-clipboard.md) 的受控边界，目标含筛选隐藏行则拒绝。清除格式复用 [41](41-fill-and-series.md) 的独立命令。

## 验证

真实 Excel：16.0 build 20527 x64，语言 2052。只操作自建文件及自有隐藏 COM 实例，宏/外链更新禁用，记录 PID、版本和哈希，采用普通打开、重算、保存重开。以下证据位于 `artifacts/compatibility`，不进入 Git。

| 证据 | 结果 |
| --- | --- |
| `xc0210-precision-excel-03` | 104 个真实 Excel `DecimalsIncrease`/`DecimalsDecrease` Ribbon 命令案例；16 种格式、正/负/零和 General 极值，保存重开稳定 |
| `xc0210-precision-probe-01` | 312 项模型/XLSX/XLSQ 往返检查，0 差异、Open XML schema 0 问题 |
| `xc0210-precision-output-01`、`xc0210-precision-verdict.json` | QuickExcel 输出的 104 格经 Excel 普通打开、保存重开后与独立期望一致；会计格式仅忽略字体相关的填充空格 |
| `xc0210-statistics-excel-01` | 三个可见性状态的求和/平均/数字计数/非空计数对照通过，固定期望进入测试夹具 |
| `xc0210-tests-04` | 95 项格式、统计、数字显示、继承、TEXT 及剪贴板相关测试通过；覆盖稀疏大选区、预算、取消、溢出和原子撤销 |
| `xc0210-wpf-03` | 18 项实际桌面处理器检查通过，含筛选统计、增减位数、格式刷/取消/过期、清除格式及保存重开；截图确认按钮和状态栏可见 |
| `xc0210-ui-oracle-01`、`xc0210-ui-output-01`、`xc0210-ui-verdict.json` | WPF 01 保存的四个货币/公式/文本编号/未修改格与独立 Excel 样本逐字段一致，普通保存重开稳定，源哈希不变 |

失败记录保留：precision-excel 01/02 暴露 PowerShell COM 的 zh-CN LCID（General 需 ClearFormats、颜色需本地化）；tests 01 修复夹具构造，02 暴露工程科学计数零值及可选小数占位差异，后续通过。WPF 01 截图暴露剪贴板第四行裁剪，已改为三行布局并补图标；02 的人工 KeyEventArgs 缺 RoutedEvent，修正烟测事件后 03 通过。

复现入口：`scripts/test-excel-format-efficiency.ps1`、`scripts/test-excel-selection-statistics.ps1`；比较工具模式 `format-efficiency <spec> <oracle> <new-directory>`；桌面模式 `--format-efficiency-smoke <directory>` 或 `scripts/test-m5-grid.ps1 -Mode FormatEfficiency -PackageDirectory artifacts/build/bin/QuickExcel.Desktop/release`。managed 命令必须指定 `--artifacts-path artifacts/build`。

Release 桌面构建通过，仅既有 NU1900 漏洞源网络警告。未运行全量回归或发布包装。WPF 路由检查不等于物理键鼠或 OS IME 验收。

参考：[CommandBars.ExecuteMso](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso)、[Excel 小数位显示](https://support.microsoft.com/en-au/excel/round-a-number)、[AGGREGATE](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.aggregate)、[SUBTOTAL](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.subtotal)。
