# XC-0202：填充、序列及清除格式

2026-10-30 完成声明范围。Ctrl+D/R、复制填充柄、数字及日/月/年序列、清除内容/格式已接入。Core 命令经 EditorController/DocumentSession 原子提交，一次撤销/重做；UI 预览不修改快照。R/D/C/E/W 沿用已有类型、公式、格式和 XLSX/XLSQ 模型，没有新增文件能力版本。

## 操作契约

`FillCellsCommand` 从原矩形沿一个方向扩展，重复原始内容和有效样式，公式复用内部剪贴板的 AST 引用平移；源矩形本身不改写。支持上下左右，空白也复制。Ctrl+D 取所选矩形首行，Ctrl+R 取首列；只选一行/一列时，分别取紧邻上一行/左一列。位于首行/首列时不修改。键盘操作保持原选区。

填充柄显示在连续矩形右下角，绿色虚线表示拟写范围；松开后提交。拖动中使用冻结窗格的模型命中地址，斜向拖动按相对扩展较大的轴选择方向。Avalonia 现已复用网格选择的边缘自动滚动，可拖出当前视口；单次源与扩展合计不超过 20,000 格。Esc、失去激活、切表/重绑取消预览；其他编辑使捕获修订过期，提交前拒绝。IME 组合、只读公式工作簿不可填充。此切片同时修复冻结两轴时编辑框错误选用不可见窗格裁剪的问题。

显式序列选中单行/单列，在“填充/清除 → 序列”指定类型和步长；首格必须为数字字面值，文本 ID、空白、公式不会自动转换。Numbers 使用 `start + index * step`，可正负/小数；日期使用整数步长，支持日、月、年。月/年始终从原始日历日计算，到目标月末截断，保留时间小数，支持 1900 虚构闰日和 1904 日期系统。日期不得为负数或超过可表示范围；1900 序号 0 只支持按日，按月/年明确拒绝。不支持工作日或增长序列。

文本编号列表现已支持显式序列：首格须为以 ASCII 数字结尾的文本（如 `XX1`、`A001`、`001`），按整数步长生成 `XX2`、`A002`、`002`；仅当源数字段有前导零时保留补零宽度，前缀原样保留。结果为文本字面值并复制首格样式。无尾随数字、步长为小数或结果溢出/为负时整批拒绝。文本 ID 始终保持文本类型。

## Avalonia 自动填充增强（2026-11-05，合并交付版本 1.15.0）

拖动填充柄默认使用 `CellFillMode.Auto`，松开时按住 Ctrl（macOS 为 Command）强制 `Copy`；Ctrl+D/R 及原有命令调用仍默认复制。悬停显示十字光标及操作提示。多行/列源按填充方向逐列/逐行独立识别，目标样式循环沿用对应源样式，源单元格不改写。

- 单个普通数字重复复制；两个及以上等差数字延续步长，支持负数、小数、向上及向左反向填充。三个及以上种子须全部符合规律；不拟合不规则数字。
- 仅数字字面值且所有种子均为日期格式时识别日期。单个日期按日递增；多个日期优先识别从首个日期出发的月/年序列，再识别整数日步长。支持两个日期系统和时间小数，月末规则复用显式序列。文本日期不解析，纯时间格式不视为日历日期。
- 单个文本尾号默认加一；多个同前缀文本尾号按整数步长延续，并验证所有种子及补零格式。负数尾号和溢出整批拒绝。普通文本、空白、混合类型、富文本或不一致模式回退复制；富文本保持原分段，公式仍按 AST 平移引用。
- 双击优先沿紧邻左列的连续数据向下填充，左侧无法延伸时尝试紧邻右列。相邻列须在源末行有数据；遇相邻列空白或目标任意列已有值/公式即停止。公式即使缓存为空也视为已有内容。没有可填充范围时显示提示，超过上限时拒绝，不静默截断。

所有填充仍走 EditorController/DocumentSession，一次撤销/重做；沿用筛选/隐藏、附件、合并、对象、富文本覆盖、工作表保护、数据验证及过期修订保护。预览与自动滚动不写入文档。尚未实现填充后模式菜单、自定义星期/月名列表、增长和工作日序列。

定向检查：`FillCommandTests`、`AutoFillTests`、`FillPlanningTests`；Avalonia `scripts/test-avalonia-workspace.ps1 -Mode Grid` 中的填充柄处理器检查覆盖自动识别、Ctrl 复制、横向文本编号、双击公式、取消、边缘滚动和一次撤销。合成鼠标事件不代表物理鼠标/IME 验收；本次自动推断契约未新增独立 Excel 对照证据。

本次 62 项定向测试通过；独立 `artifacts/fill-handle-build` 构建的 `--grid-smoke` 共 334 项通过，证据位于 `artifacts/avalonia/fill-handle/20261105-055044/result.json`。`20261105-055006` 保留 Esc 被剪切提示抢先处理的失败证据；修复为拖动取消优先后重跑通过。没有运行全量回归。

`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 已发布并激活 `artifacts/avalonia/releases/20261105-055127-b78f494f/application`；生产自检通过，`active-windows.json` 指向该包，清单及功能程序集哈希核对通过。版本 1.15.0 包含工作区同期功能改动，填充柄没有覆盖其他聊天的修改。

清除内容保持有效格式；`ClearCellFormatsCommand` 重置为工作簿默认样式 0，保留值、公式、附件和显式富文本分段字体，且不触发重新计算。真实 Excel 的 ClearFormats 也保留显式分段字体；这与显式转换纯文本是不同操作。

复制填充遇隐藏/筛选行、隐藏列、相交合并/图片或源/目标链接/批注时整批拒绝，不静默跳行。序列和清除格式保留目标附件，但仍拒绝隐藏、合并/图片目标；格式清除不会删除分段字体。内容覆盖富文本、未知函数引用变换及越界范围继续执行 [内部剪贴板](40-internal-clipboard.md) 的保护。冻结只改变视图。普通保存、未知部件损失提示、缓存新鲜度和脚本授权保持既有契约。

## 真实 Excel 对照及测试

本机 Windows x64 Excel 16.0 build 20527、语言 2052。独立 Excel 先执行 FillDown/FillRight、AutoFill 复制、DataSeries、ClearContents/ClearFormats，保存并重新打开；QuickExcel 对同一来源执行命令、撤销/重做、计算、XLSX/XLSQ 重开；再由 Excel 打开 QuickExcel 输出、重算、保存重开。捕获值、类型、公式、显示文本、格式、字体及文本逐字符字体；最终判定绑定来源和输出哈希，仅已识别 AST 序列化差异允许等价。

| 证据（均位于 artifacts/compatibility） | 结果 |
| --- | --- |
| `xc0202-excel-oracle-verified` | 17 独立来源及保存重开稳定性：方向复制、填充柄复制、清除两模式、正负/小数数字、1900/1904 日期、月末/闰日/倒序/时间小数 |
| `xc0202-probe-verified` | 17 路径、11,067 项模型/重开检查通过，撤销/重做通过，Open XML schema 0 错误 |
| `xc0202-excel-verify-verified`、`xc0202-verdict-verified.json` | 17/17，0 意外差异，包括富文本清除后的逐字符字体 |
| `xc0202-excel-selection-01` | Excel 自有隐藏实例的 CommandBars.ExecuteMso 验证单格、单行及单列的相邻来源语义，共 4 例；不是物理快捷键测试 |
| `xc0202-core-final` | 71/71 填充、剪贴板、富文本相关测试；类型、方向、引用、日期边界、撤销、拒绝、XLSX/XLSQ 往返 |
| `xc0202-rendering-final` | 138/138 WPF 处理器检查，含此前渲染/保存/剪贴板及本项初始 18 项 |
| `xc0202-fill-selection-final` | 最新填充专项 23/23；新增单格/单行/单列、首行/首列边界；冻结编辑框裁剪、填充柄与虚线预览、取消/过期、实际序列对话框、隐藏/筛选、IME 门禁、保存重开 |

复现：`scripts/test-excel-fill.ps1 -OutputDirectory <新 artifacts 目录>`；工具 `fill <oracle.json> <新目录>`；脚本 `-VerifyDirectory <模型输出目录>`；工具 `verify-fill <oracle.json> <Excel复验.json> <模型报告.json> <新verdict.json>`。单行/列命令另用 `scripts/test-excel-fill-selection.ps1`。Desktop `--fill-smoke <隔离目录>` 只运行该项 UI 专项。全部 managed 构建须指定 `--artifacts-path`。

真实 Excel 来源带有未建模的修订/拼音元数据，17 条修改路径均先验证普通保存被阻止，再显式另存诊断副本；原文件不变。不能据此宣称未知 XML 无损支持。WPF 新建簿另证正常 XLSX 保存。Desktop Release 构建通过，只有既有 NU1900 漏洞源连接警告；未跑全量回归、未发布新包。WPF 测试调用实际处理器和视口映射，不能替代物理拖动或真实 OS IME 验收。

失败证据保留：`xc0202-*-final` 中早期富文本 ClearFormats 对照失败，修正后使用 `*-verified` 通过记录；`xc0202-fill-wpf-01/02/03` 暴露冻结窗格裁剪，`fill-only-01` 是预期过期异常的测试等待处理问题，后续专项均另存证据。未以删除失败案例或放松判定获得通过。

API 依据：[Range.FillDown](https://learn.microsoft.com/en-us/office/vba/api/excel.range.filldown)、[Range.DataSeries](https://learn.microsoft.com/en-us/office/vba/api/excel.range.dataseries)、[XlDataSeriesType](https://learn.microsoft.com/en-us/office/vba/api/excel.xldataseriestype)、[XlDataSeriesDate](https://learn.microsoft.com/en-us/office/vba/api/excel.xldataseriesdate)。具体日期和清除行为以本机独立 Excel 证据为准。
