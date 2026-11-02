# QuickExcel 0.6 交互预览版

2026-11-02 Skia 候选版：59 号计划 P0–P4 声明范围完成。主网格、自动尺寸与报表使用共享文字/图片/图表渲染；打印预览可直接导出带文字和矢量的 PDF，无需安装 PDF 打印驱动。Windows/Linux 无界面输出、功能专项及原有滚动预算通过。当前候选包在 `artifacts/QuickExcel`，完整证据、测量修正与外部门禁见 [实施记录](docs/60-skia-rendering-and-pdf-implementation.md)。

XC-0201 已接入工作簿内复制/剪切及选择性粘贴，Ctrl+C/X/V 保留类型、格式并调整公式引用，可整批撤销；开始菜单提供只贴值/公式/格式和转置。真实 Excel 对照与限制见 [内部剪贴板](docs/40-internal-clipboard.md)。XC-0202 已补齐 Ctrl+D/R、填充柄复制、数字/日期序列和清除格式，真实 Excel 对照及限制见 [填充记录](docs/41-fill-and-series.md)。XC-0203 已补齐自动求和、函数补全/提示、同表点选引用、F4 和数据边界导航，见 [录入记录](docs/42-formula-entry-and-navigation.md)。XC-0204 工作表复制/重排、名称创建/改名/修改/删除已完成，见 [名称管理记录](docs/43-sheet-and-name-management.md)。XC-0205 七个条件统计函数及统一条件匹配已完成，见 [条件统计记录](docs/44-conditional-functions.md)。XC-0206 查找公式标量模式已完成，见 [查找记录](docs/45-lookup-functions.md)。XC-0207 日期/数学及工作日完成，见 [日期与数学记录](docs/46-date-and-math-functions.md)。XC-0208 文本和类型公式完成，含 Unicode 兼容版本和 #CALC! 富值读写，见 [文本与类型记录](docs/47-text-and-type-functions.md)。下一项为公式边界一致性。

XC-0107 已完成常见富文本、超链接和传统批注的读显写；右键显式转换富文本后可普通编辑，支持撤销。真实 Excel 对照及限制见 [内容记录](docs/36-xlsx-rich-text-links-and-notes.md)。[E0 证据及诊断收尾](docs/39-xlsx-diagnostic-closure.md) 已完成。

XC-0106 已完成受控公式保留：数组/结构化引用等以整簿只读打开，未知函数及依赖保留旧缓存，保存不冒充已重算。真实 Excel 对照及范围见 [公式保留](docs/35-xlsx-preserved-formulas.md)。

2026-10-30 进度同步：当前为 M5 桌面交互收敛，M1–M3 受支持子集已交付；M4 严格策略/完整部署及 M6 恢复/发布门禁仍未完成。总进度与后续顺序以 [实施计划](docs/05-implementation-plan.md) 为准。

当前优先推进 **XLSX 兼容性与 Excel 常用功能对齐**，按文件保真→高频编辑/公式→结构化数据→报表/高级分析分阶段，具体待完成清单见 [专项计划](docs/26-excel-compatibility-roadmap.md)。以下使用说明只描述已有能力。

2026-10-29 筛选排序更新：选含标题的矩形 → 数据/筛选，从数据标题格的按钮或 Alt+↓ 打开面板；支持候选搜索、多列条件、日期区间和清除条件。数据/多级排序支持多键、大小写、自定义文本序列及显式颜色；筛选时排序完整区域，包含隐藏记录，然后重新应用筛选，可一次撤销。单次排序仍限 20,000 格，筛选状态仅本次会话有效。

主工作区直接使用 `QuickExcel.Grid.Wpf`，旧 DataGrid 与“试验网格”窗口已移除；编辑、尺寸调整、筛选冻结、字体格式、合并和图片均在主界面操作。**横向拖动导致整桌面停顿已由用户于 2026-10-30 确认解决并关闭。** 默认 WPF 自动渲染；软件渲染仅作显式诊断选项。具体行为、既有证据和剩余真实 IME / 跨屏 DPI / 辅助功能边界见 [主网格迁移记录](docs/24-main-grid-migration.md)。

使用当前源码：Release 构建后运行 `artifacts/build/bin/QuickExcel.Desktop/release/QuickExcel.exe`。`Start-QuickExcel.cmd` 当前从源码运行，未指定配置时为 Debug，需要 SDK。`artifacts/QuickExcel` 是独立发布快照，不会自动更新；身份以其 `BUILD-METADATA.json` 为准。

脚本页现支持数字签名：导入 PFX/P12 或创建自签证书，保存后随 XLSQ 保留。有效且受信任的证书默认自动启用脚本；首次遇到自签证书会询问信任，可选仅本次或始终信任，也可撤销信任。修改脚本会移除签名。使用及边界见 [脚本数字签名](docs/20-script-signatures.md)；此策略更新下文及历史记录中“打开一律不启用脚本”的描述。

最新 `0.6.0-preview.3` 使用正式界面基础：启动页、默认底部的展开工作表标签（设置可放左/右侧）、独立脚本页。Ctrl+F/H 查找替换，Ctrl+G 定位，F6 切换编辑区域，Ctrl+PageUp/Down 切表。脚本页可编辑所有包内模块，新建工作簿使用 API v2。文件菜单集中兼容报告和两种导出模式。实现、图片/合并支持范围和验收见 [正式界面记录](docs/19-m5-product-workspace.md)。以下保留前轮结构与工作区背景。

结构菜单操作选区所在整行/整列；筛选选择含表头的矩形，排序使用明确的数据区域或选区。B2 冻结第一行/列，冻结行列由同一网格绘制。格式菜单修改样式，可一次撤销，原始输入不变。前轮实现与历史边界见 [结构与工作区](docs/18-m5-structure-and-workspace.md)。

文件菜单提供最近 10 文件、清空记录和关闭工作簿（Ctrl+W）；脚本通过右上角独立页面进入。保存发现外部修改/删除时保留编辑并阻止覆盖，请另存新路径或重新打开。筛选/冻结仅本次工作区有效，重开复位；筛选下复制可见行、粘贴跨筛选隐藏行拒绝、Delete 只清可见格。已声明数据区域可在筛选状态下排序全部记录，再重新应用筛选。

2026-10-28 的 [M5 首切片记录](docs/17-m5-grid-and-input.md) 保留旧 DataGrid 的中文拼音、编辑器焦点及保存重开实机证据；不直接作为新 FastGrid 的完整输入验收。

2026-10-28：已补齐宿主崩溃后的 profile/payload 回收，625 项测试通过。本机部署矩阵五项通过；直接长路径 EXE 启动失败仍待处理，详见 [回收与部署验证](docs/16-m4-recovery-and-deployment.md)。

本版保留 M2/M3 文件与计算及 M4 worker/模块/actions 能力，正在收敛 M5 桌面交互。发布包需要 Windows x64 与 .NET 10 Windows Desktop Runtime；保留完整发布目录后运行 QuickExcel.exe。

最新切片已把无能力 AppContainer 接入公式与 actions 的实际 worker，并补齐回环/DNS 的本机 WFP 拒绝证据。详见 [隔离 ADR](docs/15-m4-os-isolation-adr.md)和 [合批及容量](docs/14-m4-batching.md)；部署矩阵、严格策略及 M6 OS 门禁仍未完成，仅供可信脚本使用。

## 使用

1. 打开或新建工作簿。未签名文件的脚本保持停用；有效且受信任的签名可在桌面验证后自动启用并重算。恢复副本不自动启用脚本；已有公式缓存可能过期。
2. 双击单元格或在公式栏编辑。前置单引号强制文本，`00123` 保留文本，单独输入单引号表示空文本。
3. Ctrl+S 保存 XLSQ v2；活动单元格与公式栏会先提交。输入无效阻止保存；重算失败仍可保存有效输入与旧缓存。
4. v1 文件迁移后显示未保存状态，首次保存选择其他文件名，旧文件保留。
5. 含未知部件的文件会显示兼容报告。未修改工作簿内容时可以原样保留；编辑后可取消，或明确另存移除这些内容的副本。重新打开副本可刷新兼容报告。
6. “导出值快照”将公式全部转为结果；“导出公式”冻结脚本及传递依赖结果，保留独立标准公式。需要冻结的结果缺失时拒绝，过期结果须明确选择。
7. 在脚本专页点击“运行”（Ctrl+Enter）应用源码并重算。源码改变会立即撤销授权。API v2 支持包内模块编辑和用途设置；API v1 保留旧语义。
8. 包含 actions 绑定的文件应用后显示动作列表。选择并运行，全部写值/公式通过校验后一次提交，Ctrl+Z 一次撤销，重做不再次运行动作。动作执行期间可停止；冲突或失败后须重新运行。动作提交后公式缓存等待重算。

## 编辑和快捷键

拖动列标题右边界调整列宽，拖动行号下边界调整行高；松开后一次提交，可撤销、重做并保存。双击边界按显示内容自动调整尺寸；Escape、窗口失活或捕获丢失取消拖动预览。冻结区域同步尺寸，菜单也可输入精确尺寸。

Ctrl+O 打开，Ctrl+N 新建，Ctrl+Shift+S 另存。单击单元格只选中，已选中单元格也可按住鼠标拖选；双击或 F2 进入编辑，直接键入替换。选中时显示计算结果和绿色边框，编辑时显示原始输入与光标。Enter/Shift+Enter 提交后向下/上，Tab/Shift+Tab 提交后向右/左，Escape 取消草稿；Shift+方向键扩展矩形选区。输入法组合期间请先确认或取消候选，再保存或切换。

Ctrl+C 复制矩形显示结果，非连续选区拒绝复制。Ctrl+V 粘贴值，`=` 开头保留为文本；Ctrl+Shift+V 显式按输入解析公式。单值填充选区，矩阵需选左上角或同尺寸矩形；支持引号包裹的多行/Tab 文本。Delete 清空实际选中位置。一次粘贴/清空可整体撤销。

Ctrl+Z/Y 在文本编辑中作用于当前文本框，退出编辑后作用于工作簿事务。默认历史预算 32 MiB，截断时显示提示。大选区按范围保存，复制最多 100,000 格、4 Mi 个字符，选择容量不等于写入容量。复制跳过手动隐藏和筛选隐藏行；区域操作可包含手动隐藏记录，矩阵粘贴仍不能跨筛选隐藏行。

一次粘贴最多 20,000 格，任一目标非法则整批不写。界面最多显示 100,000 行 × 256 列；模型最多 64 表、1,024 列、全簿 1,000,000 存储格，文件还须满足解压与解析预算。远端空洞不展开为密集矩形。

## 文件与脚本

v2 保存唯一的 workbook.xlsx，以及身份/修订状态和脚本源码模块映射。支持数字、布尔、标准错误、文本、空白/空文本、公式缓存、1900/1904 日期系统及基础格式。历史 v1 中已经丢失的类型不猜测恢复。

```javascript
quickExcel.functions.register("DISCOUNT", (price, rate) => price * (1 - rate));
quickExcel.functions.register("DOUBLE", value => value * 2);
```

支持同步标量、v1 一维区域与 v2 显式二维区域参数；没有 Promise、任意对象或动态数组溢出支持。QuickJS 每轮运行于 AppContainer worker，执行期间可点击状态栏「停止计算」。文件级 API v2 支持按清单逻辑 ID 导入 ES 模块、多个入口和共享依赖；初始化必须同步完成。actions 使用独立上下文、只读快照和有界 patch，首版每次最多写 20,000 格；含合并或单元格内图片的工作表暂不接受动作写入。打开、迁移、另存均不会复制运行授权。OS 最终发布验收仍待完成，继续仅供可信脚本使用。发布包的 `ScriptWorker` 目录必须保留；见 [actions 示例与边界](docs/13-m4-actions.md)、[模块契约](docs/12-m4-modules.md)及 [worker 支持范围](docs/11-m4-worker.md)。

跨表计算已接入；未授权脚本和依赖仍保留 Blocked 缓存。主题颜色已解析，主题字体/效果/关联、富文本、图表/透视表、复杂锚点、单元格内图片与数组公式等仍有保真限制；见 [M2 文件与类型](docs/09-m2-files-and-types.md)及 [XLSX 渲染](docs/25-xlsx-rendering.md)。基础图片/合并、排序筛选已接入。恢复目前只有手动打开 `.xlsq` 副本，自动恢复写入/发现及完整无障碍矩阵仍待完成。

## 源码目录

Desktop 按界面职责组织，根目录保留 `App.xaml` / `App.xaml.cs`、项目文件、说明和许可证。

```text
QuickExcel.Desktop/
├─ Composition/          具体存储、计算和脚本运行时的依赖组装
├─ Controls/
│  ├─ Grid/              FastGrid 适配、自动尺寸与绘制
│  └─ Ribbon/            Ribbon 附加属性
├─ Views/
│  ├─ MainWindow/        主窗口 XAML、code-behind 和各功能 partial
│  └─ Dialogs/           排序对话框、筛选弹层
├─ Editing/              WPF 原生文本编辑器辅助
├─ ViewModels/           文档状态的只读视图模型
├─ Themes/               应用共享的 XAML 样式与模板
├─ Diagnostics/
│  ├─ Smoke/             WPF smoke 场景、回归检查及共用验证辅助
│  └─ Benchmarks/        网格、拖动、滚动和尺寸调整的性能诊断
└─ docs/                 产品、架构、契约与验收记录
```

新增主窗口功能放入 `Views/MainWindow/MainWindow.<功能>.cs`；独立控件和对话框分别放入 `Controls` 与 `Views/Dialogs`。诊断场景放入 `Diagnostics`，避免继续堆积在界面实现旁。共享规则和编辑控制器已迁入 `QuickExcel.Presentation`，尺寸/范围算法迁入 `QuickExcel.Grid.Shared`；Core.Test 使用项目引用，不再链接这些源文件。所有文档修改仍经共享 `EditorController` 和 `DocumentSession`。

这次整理只调整物理路径，保留 `QuickExcel.Desktop` 命名空间和现有类型；主窗口诊断仍是同一 `MainWindow` 的 partial，继续随 Desktop 编译并通过原命令行入口运行。SDK 自动包含子目录中的 C# / XAML 文件，无需逐文件维护项目项。`App` 仍直接创建窗口，主题资源路径保持 `Themes/Ribbon.xaml`。

## 构建与验收

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\QuickExcel -Mode All
```

从仓库根目录执行。所有引用项目输出也留在本仓库 artifacts。上述完整检查用于基线/发布验收；日常变更只运行相关测试与 WPF 专项。验证当前源码时将 `-PackageDirectory` 指向最新 Release 构建目录，避免误用旧发布包。发布脚本生成 ZIP、SHA-256 和 BUILD-METADATA；无 SDK 干净机器、程序/安装包签名或安装器验收仍未完成。详见[发布说明](RELEASE-NOTES.md)。

计算范围与函数列表见 [M3 支持矩阵](docs/10-m3-calculation.md)。新计算器拒绝脚本覆盖内置名称；请使用 JS 命名空间。结构命令和当前范围的界面操作已接入，剩余 M5 验收以实施计划为准。
