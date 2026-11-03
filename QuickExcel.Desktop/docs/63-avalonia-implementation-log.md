# Avalonia 迁移实施日志

2026-11-03。计划见 [61 号文档](61-avalonia-migration-plan.md)，A0 的依赖、运行与性能基线见 [62 号记录](62-avalonia-a0-validation.md)。阶段提交已在 `codex/avalonia-migration` 分支推送，用户已同意按 Windows 预览版范围并入 main。当前默认入口为 Avalonia，A7 已移除 WPF 生产项目；平台限制及合并前核验见 [68 号记录](68-avalonia-retirement.md)。以下阶段结果描述各自受测源码，保留历史证据。

## 验收顺序调整

用户要求：“涉及实机验证的先跳过，在实施日志中记录即可，后期统一验证。”因此实机项延期，不阻塞工程实施，也不标记为通过。Linux 虚拟机用户实测无法输入中文、无候选框，保留为已知未解决问题。macOS、真实打印、跨屏 DPI、辅助技术、干净设备部署与真实设备性能继续待验收。不得用此调整降低脚本隔离、文件保真或草稿安全要求。

## A1 共享层（完成）

- 新建 `QuickExcel.Presentation`（net10.0）：迁入 11 个 Interaction 源文件、EditorController、IGridSurface；提供公共共享类型。编辑、验证确认、预期修订、撤销事务、公式粘贴、组合门禁与原单元格草稿所有权逻辑保持。
- 新建 `QuickExcel.Grid.Shared`（net10.0）：移动 SeriesSizes、FastGridCellAddress、ActiveSeries 和矩形/合并选区算法。保留历史命名空间、原始来源哈希与 MIT 声明；WPF 与测试均引用同一程序集，不链接重复源。许可发布到独立子目录，避免两个项目的同名输出冲突。
- 新建本地文件选择端口，明确取消返回 null、非本地存储 URI 拒绝。WPF 打开/保存使用适配器，现有注入对话框和 SaveChecked/打开指纹流程保留；没有把所有 URI 当成本地路径。
- Core、Infrastructure、计算、文件格式及 worker 实现未改动；共享程序集不引用 WPF、Avalonia、Infrastructure 或 native。

实际验证：

| 检查 | 结果/证据 |
| --- | --- |
| Release 解决方案构建 | 0 警告、0 错误；输出 `artifacts/build` |
| 相关规则与新契约测试 | **177/177**；`artifacts/avalonia/tests/a1-shared-final.trx` |
| WPF 发布与交互 | `artifacts/avalonia/a1-wpf`；**95** 项检查，`artifacts/validation/m5-grid/20261102-164457-15d1682c` |
| WPF 工作区 | **157** 项检查，`artifacts/validation/m5-grid/20261102-164503-b7c96f15`；包含关闭取消、保存与冲突等原入口 |

定向测试覆盖编辑原子历史、确认期间修订冲突、端口路径限制、共享程序集依赖、剪贴板/组合规则、尺寸/超大格/合并选区、筛选排序、查找与兼容报告。首次执行暴露两项旧测试仍硬编码 100,000 行和 4,096 样式（现有限额已是 200,000/8,192）；改用 `WorkbookLimits` 验证真实边界后通过，未调整产品限额。首次发布的同名许可证冲突已修复并重发。WPF smoke 为路由事件/合成交互证据，不冒充物理输入。没有运行全量 Core 回归或重复性能基准。

## A2 Avalonia 工作区（完成）

新建 `QuickExcel.Desktop.Avalonia` 与 `QuickExcel.Grid.Avalonia`，加入解决方案与精确锁文件。Fluent 主题使用开源基础控件；启动页、最近文件、新建/示例、打开、保存/另存、撤销/重做、基本工作表管理、缩放/标签位置、公式栏及关闭等待已接入。文件选择验证本地路径；共享 WorkbookWorkspace 在打开成功后替换文档，SaveChecked 使用打开/上次保存的指纹，IO 失败保留编辑及绑定。

基础网格复用 SeriesSizes，新增可独立测试的 ViewportAxis；只枚举可见轴，绘制有 1,600 万像素预算的独立不可变位图，选区单独绘制，一个原生 TextBox 承接当前格。原格草稿、取消、预编辑阻止提交和隐藏公式/只读限制已经接入。内容经共享 SkiaRenderer，条件格式投影由宿主注入，Grid.Avalonia 不引用 Infrastructure。Native TextBox 的模板与 PreeditText 由 CompositionTextBox 观察，未替换系统 IME client。

Windows 组合根保持原 ScriptWorker 路径，非 Windows 明确关闭脚本，未新增弱隔离后备。没有切换默认启动器。

| 检查 | 结果/证据 |
| --- | --- |
| 最终解决方案 | 19 个项目构建，0 警告、0 错误，WPF 基线仍可构建 |
| 几何/既有滚动/超大格/共享契约 | **21/21**，`artifacts/avalonia/tests/a2-geometry-final.trx` |
| 共享文件会话/几何/共享契约 | **12/12**（与上一行部分重叠），`artifacts/avalonia/tests/a2-workspace.trx` |
| Avalonia 工作区 | **32** 项通过，进程退出 0；`artifacts/avalonia/workspace/20261102-170706-1c4de943`，每次独立目录 |
| 截图复核 | 修复 DPI 二次放大后，内容/选区同坐标；原生公式栏可见；截图在各次 `workspace.png` |
| 发布与载荷扫描 | `artifacts/avalonia/a2-package`；无 WPF、WorkerFixture、SandboxHostFixture、SandboxProbe 文件；完整 ScriptWorker/桥 DLL/许可证存在 |

首次截图发现位图源矩形使用 DIP 导致高 DPI 裁剪并重复放大，改用实际像素源矩形；派生 TextBox 未应用基础主题导致模板缺失，显式继承 TextBox style key 后增加真实 presenter 与合成 preedit 检查。几何测试纠正视口最后一个部分可见格的期望。ViewportAxis 使用 `QuickExcel.GridGeometry` 命名空间，避免共享程序集引入的 `QuickExcel.Grid` 与 WPF Grid 类型名称冲突。以上失败与修复分别保留在各次 artifacts 记录。

工作区检查包括公式栏原格草稿、取消编辑、关闭取消、保存对话框取消、中文路径、重算、撤销/重做、重开、导航不标 dirty、持久缩放、设置隔离、打开失败、外部修改冲突及另存恢复。独立文件会话测试增加目标目录不存在、源文件删除和恢复解绑。未用这些结果关闭物理输入、辅助技术或性能门禁。

## A3 网格与非 Web 功能对齐（工程完成）

已完成 AV-021–AV-024 的原生网格与非 Web 工作区，详细入口矩阵、生命周期、失败记录及验收边界见 [64 号 A3 记录](64-avalonia-grid-and-workspace.md)。异步数据验证保留原格与修订；文件导出不改变主文件绑定/历史；所有写入继续走共享命令。原生功能区、公式辅助、剪贴板、填充、格式/结构/搜索、数据工具、图片/图表/批注、页面元数据、保护、恢复与有界无障碍信息均已接入。默认完整坐标域可达，不再需要旧网格的“显示更多行/列”。

实际验证：19 项目 Release 构建 0 警告/0 错误；105 项相关共享测试；77 项 Avalonia 网格/入口检查；32 项 Avalonia 工作区检查；157 项 WPF 工作区回归。原始目录在 64 号记录中逐项登记，没有跑全量 Core 套件。首轮横向普通/带样式 p95 为 11.5642/9.9682 ms，拖选间隔为 19.0615 ms；最终候选分别为 16.1515/9.1766/19.2976 ms，两次均达到原预算且数据全部保留。25,000/100,000/250,000 存储格、8,192 样式和远端坐标通过有界绘制负载。

前期失败与修正保留：公式标准化括号断言、普通文本隐藏目标拒绝、函数补全排序、误用未建模 ROW 的测试、未绑定网格布局、主滚动条识别和导出测试前置格式。没有改变公式语义或扩大已支持格式。共享 SetAxisMetadataCommand 增加稀疏“显示全部”处理，避免生成默认行元数据；原 WPF 回归通过。

实机项按用户要求延期，Linux 中文输入仍未解决；CPU 绘制提交/合成事件不冒充真实显示或输入法证据。A3 完成时，A4 的 Web/脚本接入和 A5–A7 尚待后续实施，默认启动器仍是 WPF；后续工程状态以本日志 A4/A5 及 67、68 号记录为准。Avalonia 设置使用独立 LocalApplicationData/QuickExcel/Avalonia 目录。

## A4 Web 组件（工程完成）

正式 Monaco/PDF.js 页面、文档/视图代次消息协议、草稿 flush/ack、隔离 worker 操作/签名流程、不可变打印会话与 Windows 原生打印代码已接入。Windows 42 项、Linux WebKit 连续两轮各 35 项通过；独立冷启动、编辑/PDF/并发性能与完整进程树资源检查通过。逐项证据、已修复问题和未完成项见 [65 号日志](65-avalonia-web-components.md)。本阶段单独提交；实际输入法、打印、DPI 与设备验收按用户要求延期，不视为通过。

## A5 Windows 切换（工程完成）

默认启动器和发布入口已切换 Avalonia，Windows x64 自包含正式候选含完整 AppContainer worker、离线 Web、来源锁与文件身份。发布不编译测试注入并排除设计器/符号/测试载荷；安装自检实际执行 Monaco、PDF.js、存储和隔离 worker。设置一次性复制、备份和原子提交；信任数据不迁移。旧 WPF 包经 A0 身份核验后归档，已实际往返切换并运行旧 157 项和新 7 项检查。

77 项定向协议/worker/恢复检查、77 项网格检查、32 项工作区检查及完整独立 Web 性能复核通过；中文/空格及只读安装副本成功，缺失资源/worker 明确拒绝。最终包、哈希、失败修复与证据目录见 [66 号记录](66-avalonia-windows-release.md)。实机验收仍延期；A6 的原生桥、隔离、信任和打印按平台独立实施，不以本阶段结果宣称跨平台完整支持。

## 2026-11-03 功能区、列操作与筛选面板修正

- Ribbon 固定为 170 DIP；页内分组横向排列、三行按钮及底部组名统一对齐，字体/字号并排。溢出可用横向滚动条或滚轮访问；切换页签、调整窗口宽度不再挤高网格上方区域。页签采用开始/插入/页面布局/公式/数据/审阅/视图顺序。
- 列筛选改为列头定位的轻量浮动 Popup，点击外部、Esc、取消可关闭。控件属于宿主逻辑树，获得完整原生主题、候选虚拟化和键盘焦点；不再建立模态 Window。排序、清除、搜索、全选/部分选择、双条件、确定/取消均保留。
- 候选模板接受空数据上下文；勾选状态归候选模型所有，不依赖被回收的 CheckBox。搜索只改变显示列表，全选只影响搜索结果。空选择、非法数字/日期和持久化条件转换错误在面板内显示。候选显示格式不受支持时仍可清除或编辑条件。异步候选加载和提交核验会话、工作表与修订；切换文档、修订改变及主窗口关闭会释放面板等待。
- 整列/整行的逐格格式和清除命令在枚举前收缩到该选区的已使用范围（含样式格、链接、批注定位），不再把整列空白容量计入 20,000 格限制；这些命令作用于当前已用范围，不建立未来空白行的默认列样式。显式矩形保持原范围。排序另按记录内容推断范围，包含偏移表头、空行之后的记录及末尾空排序键的相邻记录，扩展相邻列以整条记录移动；Table 保留表头和汇总行。单列文本表头加数值数据可识别表头，纯文本区域仍由多级排序指定。真正超限的排序拒绝，不静默缩成小区域；筛选范围推断不继承排序容量限制。

验证：Release 解决方案构建 0 警告/0 错误；`SortExpansionTests|PersistentFilteringTests|M5WorkspaceTests` **55/55**；新增 `scripts/test-avalonia-workspace.ps1 -Mode RibbonFilter` **41 项**，包含反复开关/搜索、条件错误、候选失败、虚拟化滚动、取消/撤销、文档失效、Table 汇总行和整列入口。证据及两张布局截图：`artifacts/avalonia/workspace/20261103-013412-9e4fe00c`。未运行全量回归，也未以程序化 UI 检查替代物理输入、DPI/辅助技术验收。

本地 Windows 包已通过 7 项生产自检并激活：artifacts/avalonia/releases/20261103-013523-ca0842a1/application；发布使用 build-quickexcel.ps1 -SkipTests -RunSmoke（相关测试已单独完成），原包保留。

字体/字号裁剪跟进：将 ComboBox 的默认 32 DIP 最小高度调整为 28 DIP，并移除功能区 ScrollViewer 的 126 DIP 固定子高度，使其适应页签内容视口；外层 Ribbon 仍固定 170 DIP。新增沿真实视觉祖先链的边界检查，复现了行内溢出与页签视口二次裁剪；修复后 RibbonFilter 65 项检查通过，截图已复核（artifacts/avalonia/workspace/20261103-014517-367e6344）。

## 2026-11-03 自动执行 Ribbon 页

在“公式”之后增加“自动执行”页，Actions 分组通过下拉框显示当前工作簿所有 actions 入口注册的名称和描述。上方关键词框按名称或描述进行包含匹配，忽略大小写和关键词两端空白；无匹配时显示提示，清空关键词恢复全部选项。单击、双击和键盘确认均只选择，只有“执行”按钮运行动作。下拉选项可纵向滚动，长描述通过悬停完整显示，无描述时显示“无描述”。功能区仍保持 170 DIP 高度。

“加载操作”复用现有“应用并重算”的授权、签名检查与隔离 worker 发现流程，无需打开脚本编辑器；从公式页或脚本编辑器应用脚本也会更新此列表。未授权工作簿打开时显示加载提示，不隐式运行模块初始化。脚本源码/绑定变化、停用、文档切换均清除旧列表；无选择、只读、运行期间或平台不支持时禁止执行。运行前提交当前编辑，动作仍通过 DocumentSession 原子提交并可一次撤销，异常显示在状态栏。

初版列表验证（后续已移除双击执行）：Avalonia Release 构建 0 警告/0 错误；`scripts/test-avalonia-workspace.ps1 -Mode Automation` **26 项**通过，覆盖真实多入口 worker 发现、按钮/双击、重复请求、单次撤销/重做、异常无部分写入、编辑草稿提交及旧列表失效；证据与截图：`artifacts/avalonia/workspace/20261103-015513-33a0d24e`。复用分组布局后，`-Mode RibbonFilter -SkipBuild` **67 项**通过：`artifacts/avalonia/workspace/20261103-015533-0629e831`。未运行全量回归；这些程序化检查不替代物理鼠标、DPI 或辅助技术验收。

下拉搜索跟进：匹配范围保留当前选择；选择被筛掉时清除选择并禁用执行，清空关键词不会暗中恢复旧选择。停用脚本或列表失效时同时清空关键词并关闭下拉框。搜索编辑和下拉导航保留自己的键盘输入，Delete/剪切不再作用于单元格，Alt+Down 不再误触工作表筛选。更新后的 Automation **62 项**通过，包含名称/描述搜索、空结果、选择失效、双击不执行、键盘选择、最小窗口高度和真实 worker 执行；Release 构建 0 警告/0 错误，证据及复核截图：`artifacts/avalonia/workspace/20261103-021053-c276eaaf`。相关 RibbonFilter **67 项**通过：`artifacts/avalonia/workspace/20261103-021130-2a26764a`。键盘检查等待延迟重算结束后再发出合成按键，避免忙碌状态影响焦点；仍未运行全量回归或进行物理输入验收。
