# Avalonia 迁移实施日志

2026-11-03。计划见 [61 号文档](61-avalonia-migration-plan.md)，A0 的依赖、运行与性能基线见 [62 号记录](62-avalonia-a0-validation.md)。每阶段独立本地提交，用户验收后才推送远程。

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

实机项按用户要求延期，Linux 中文输入仍未解决；CPU 绘制提交/合成事件不冒充真实显示或输入法证据。A4 接入 PDF.js、Monaco 与脚本授权/签名/日志；A5–A7 尚未开始。当前设置使用独立 LocalApplicationData/QuickExcel/Avalonia 目录，默认启动器仍是 WPF。

## A4 Web 组件（工程完成）

正式 Monaco/PDF.js 页面、文档/视图代次消息协议、草稿 flush/ack、隔离 worker 操作/签名流程、不可变打印会话与 Windows 原生打印代码已接入。Windows 42 项、Linux WebKit 连续两轮各 35 项通过；独立冷启动、编辑/PDF/并发性能与完整进程树资源检查通过。逐项证据、已修复问题和未完成项见 [65 号日志](65-avalonia-web-components.md)。本阶段单独提交；实际输入法、打印、DPI 与设备验收按用户要求延期，不视为通过。

## A5 Windows 切换（工程完成）

默认启动器和发布入口已切换 Avalonia，Windows x64 自包含正式候选含完整 AppContainer worker、离线 Web、来源锁与文件身份。发布不编译测试注入并排除设计器/符号/测试载荷；安装自检实际执行 Monaco、PDF.js、存储和隔离 worker。设置一次性复制、备份和原子提交；信任数据不迁移。旧 WPF 包经 A0 身份核验后归档，已实际往返切换并运行旧 157 项和新 7 项检查。

77 项定向协议/worker/恢复检查、77 项网格检查、32 项工作区检查及完整独立 Web 性能复核通过；中文/空格及只读安装副本成功，缺失资源/worker 明确拒绝。最终包、哈希、失败修复与证据目录见 [66 号记录](66-avalonia-windows-release.md)。实机验收仍延期；A6 的原生桥、隔离、信任和打印按平台独立实施，不以本阶段结果宣称跨平台完整支持。
