# QuickExcel 架构与实施文档

2026-11-03：新增 [JSC 单引擎替换与 TypeScript 接入方案](70-jsc-typescript-migration-plan.md)，覆盖引擎嵌入、自动编译、文件与授权、隔离、分阶段迁移及验收。状态为目标方案，尚未实施；当前生产引擎仍为 QuickJS-NG。

2026-11-03：[Avalonia 迁移](61-avalonia-migration-plan.md) **A0–A5 工程完成，A6 实施中**。默认入口已切换 Avalonia，Windows 发布与 WPF 回退见 [66 号记录](66-avalonia-windows-release.md)。实机项目按用户要求延期；Linux 中文输入无候选框仍未解决，非 Windows 完整支持尚未完成。

2026-11-02：[SkiaSharp 方案](59-skia-rendering-and-pdf-plan.md) **P0–P4 声明范围完成**，包括共享文字/图片/图表渲染、直接 PDF、按需页面预览、主网格和旧依赖清理。Windows/Linux 实际输出、功能定向验证和原有滚动预算通过，完整状态、测量修正与外部门禁见 [60 号实施记录](60-skia-rendering-and-pdf-implementation.md)。

2026-11-02 当前状态：XLSX / Excel 对齐计划 E0–E4 已逐项完成声明范围；[E4 报表、打印与容量](57-e4-reporting-and-capacity.md) 的五项实现、289 项相关回归、最终 117 项专项、81 项 Excel、23 项 WPF、PDF/DPI 和容量分级验收完成。以下较早的阶段描述为历史；E5、M4/M6 剩余门禁未完成，本轮未发布新包。

XC-0001/0002 完成：[真实 Excel 业务基线](37-excel-business-baseline.md) 覆盖五类六文件，5384 属性通过；[独立编辑矩阵](38-excel-edit-matrix.md) 的 48 路径及 Excel 重开对照通过，原始差异、诊断副本和预期公式表示均明确登记。[XC-0003 诊断收尾](39-xlsx-diagnostic-closure.md) 完成，E0/E1 按声明范围关闭，[XC-0201 内部剪贴板](40-internal-clipboard.md)、[XC-0202 填充与序列](41-fill-and-series.md) 已完成声明范围，[XC-0203 录入和公式编辑](42-formula-entry-and-navigation.md) 完成声明范围，[XC-0204 工作表与名称管理](43-sheet-and-name-management.md) 已完成声明范围，[XC-0205 条件统计公式](44-conditional-functions.md) 已完成声明范围，[XC-0206 查找公式](45-lookup-functions.md) 已完成声明范围，[XC-0207 日期/数学及工作日](46-date-and-math-functions.md) 已完成声明范围，接续 XC-0208。

XC-0107 已完成常见富文本、超链接和传统批注的读显写；右键显式转换富文本后可普通编辑，支持撤销。真实 Excel 对照及限制见 [内容记录](36-xlsx-rich-text-links-and-notes.md)。E0 证据现已补齐，下一步进入 E2。

XC-0106 已完成受控公式保留：数组/结构化引用等以整簿只读打开，未知函数及依赖保留旧缓存，保存不冒充已重算。真实 Excel 对照及范围见 [公式保留](35-xlsx-preserved-formulas.md)。

**当前进度入口：[实施计划](05-implementation-plan.md)（2026-11-02 同步）。** M1–M3 受支持子集已交付；当前为 M5 桌面交互收敛，M4 严格策略/完整部署与 M6 恢复/性能/发布仍未完成。横向拖动导致整桌面停顿已由用户确认解决，关闭记录见 [24 号文档](24-main-grid-migration.md)。

**当前重点工程：[XLSX 兼容性与 Excel 功能对齐计划](26-excel-compatibility-roadmap.md)。** 2026-10-30 按用户要求提升为当前优先工作，提供 E0–E5 阶段、已交付基线和待完成清单，逐项标记重要性、常用程度、复杂度及验收条件。

阅读时先看实施计划的当前状态和下一轮切片，再查专题契约与证据。日期早于当前状态的“下一步”“默认网格”“尚未交付”属于当时记录，不覆盖新状态。Avalonia 源码版本为 `0.6.0-preview.5`；旧发布目录不会随构建自动更新，身份以包内 `release.json` 为准，旧 WPF 包保留 `BUILD-METADATA.json`。

E0 的 [兼容基线](27-xlsx-compatibility-baseline.md) 包含当前 R/D/C/E/W 矩阵、12 个合成样本和独立校验；[真实 Excel 样本台账](37-excel-business-baseline.md) 已补齐。编辑后预期差异和诊断定位已分别按 38/39 号记录完成，E0 声明范围关闭。

XC-0003 首切片已接入：[兼容诊断分级](28-xlsx-compatibility-diagnostics.md) 记录五类诊断、来源/当前单元格定位、缓存说明及保存保护回归。

E1 七项实现及各自真实 Excel 对照已完成；[行列样式优先级](29-xlsx-row-style-inheritance.md) 保留首切片历史，最新契约见 30–36 号专题记录。

新增 [XLSX 渲染修复](25-xlsx-rendering.md)：行高单位、字体行框裁剪、主题色、宽列与滚动定位，以及兼容性报告和附件复现证据。

新增 [主网格迁移记录](24-main-grid-migration.md)：主界面已切换到 `QuickExcel.Grid.Wpf`，旧 DataGrid 与试验窗口已移除。编辑、尺寸、冻结、筛选、样式和图片迁移有专项证据；真实 IME、跨屏 DPI、辅助功能及完整性能门禁仍待验收。[23 号记录](23-fast-grid-implementation.md) 保留试验阶段和历史性能对照。

新增 [类 Excel 功能区](21-excel-ribbon.md)：选项卡、分组、FluentIcons 图标、字体与对齐控制、窄窗口行为及专项验证。

新增 [工作簿脚本数字签名与证书信任](20-script-signatures.md)：签名、自签证书、自动授权和兼容性；此策略更新历史文档中“打开一律不执行脚本”的描述。

`0.6.0-preview.3` 的正式界面基础见 [M5 正式界面与桌面流程](19-m5-product-workspace.md)：启动、导航、脚本专页、查找替换、基础图片/合并，以及当时验证范围。后续主网格迁移和修复以 24/25 号记录为准。

设计基线：2026-10-22；17/18 号文档保留旧网格的 M5 首切片与结构工作区证据。新主网格真实 IME、跨屏 DPI、辅助功能仍待验收；自动恢复目前只有手动打开副本入口。M4 部署和隔离边界继续见 [16](16-m4-recovery-and-deployment.md) 与 [15](15-m4-os-isolation-adr.md)。

当前使用 XLSQ v2、类型化跨表计算及独立 AppContainer worker。Core 保持纯 .NET；自有存储/计算与脚本适配位于 Infrastructure，已移除 ExcelIO 和 ExcelIO.Formula 依赖，见 [清理记录](58-excelio-removal.md)。下文明确标记 0.2 的内容保留为历史。

## 阅读顺序

| 文档 | 解决的问题 |
| --- | --- |
| [JSC 单引擎替换与 TypeScript 接入方案](70-jsc-typescript-migration-plan.md) | JSC 原生桥、TS 自动处理、缓存/签名/IPC、平台隔离、J0–J6 迁移与定向验收；尚未实施 |
| [Avalonia 迁移方案与实施计划](61-avalonia-migration-plan.md) | MIT 依赖准入、模块边界、主网格迁移、PDF.js/Monaco、A0–A7 工作项、平台隔离与发布回退；A0–A5 工程完成 |
| [Avalonia 实施日志](63-avalonia-implementation-log.md) | 分阶段实现、验证与延期事项 |
| [Avalonia 网格与工作区](64-avalonia-grid-and-workspace.md) | 功能入口矩阵、交互及滚动性能 |
| [Avalonia Web 组件](65-avalonia-web-components.md) | Monaco/PDF.js、协议、打印会话与浏览器证据 |
| [Avalonia Windows 发布](66-avalonia-windows-release.md) | 正式包、依赖与设置、切换回退及部署证据 |
| [Avalonia A0 技术验证与基线](62-avalonia-a0-validation.md) | 固定来源/锁文件、Windows/Linux 探针、WebView 生命周期修复、离线 worker/虚拟 HTTPS、WPF 基线及待验收门禁 |
| [SkiaSharp 引入方案](59-skia-rendering-and-pdf-plan.md) | 原始跨平台渲染、PDF 生成/预览、分阶段迁移方案；P0–P4 已实施 |
| [Skia 实施记录](60-skia-rendering-and-pdf-implementation.md) | 实际架构/预算/依赖、Windows/Linux 输出、原有性能预算、候选包证据与外部门禁 |
| [分组与分类汇总](56-outline-and-subtotals.md) | E3 收尾：outline、SUBTOTAL、预览生成、真实 Excel 验收及混合状态边界 |
| [XLSX / Excel 对齐专项](26-excel-compatibility-roadmap.md) | 当前重点工程的阶段、优先级、支持维度及完成清单 |
| [XLSX 兼容基线](27-xlsx-compatibility-baseline.md) | 当前特性矩阵、样本台账、独立比较工具、首轮缺陷与证据 |
| [XLSX 兼容诊断](28-xlsx-compatibility-diagnostics.md) | 分级报告、未知函数与旧缓存、拒绝定位及保存保护证据 |
| [XLSX 行列样式](29-xlsx-row-style-inheritance.md) | 行格式应用标志、单格/行/列优先级、清除内容和读改存证据 |
| [Excel 数字格式](30-excel-number-formats.md) | 小数、货币/会计、分段/颜色、科学计数及真实 Excel 对照 |
| [命名样式与主题](31-xlsx-named-styles-and-theme.md) | XF/命名关联、主题保留、字体回退、旋转/缩进/缩小填充 |
| [XLSX 保存与文档属性](32-xlsx-save-and-metadata.md) | 公式保真保存、属性白名单、依赖判定、脚本主文件与冲突保护 |
| [实施计划](05-implementation-plan.md) | 当前 M1–M6 状态、已关闭缺陷、后续顺序与退出条件；进度统一入口 |
| [主网格迁移](24-main-grid-migration.md) | 当前主界面接线、旧控件移除、功能回归、性能样本及剩余验收边界 |
| [FastWpfGrid 实施记录](23-fast-grid-implementation.md) | 使用入口、筛选/排序支持矩阵、专项测试、性能证据和候选迁移缺口 |
| [FastWpfGrid 增强与筛选排序方案](22-fastwpfgrid-and-data-operations.md) | 当前实现状态、目标支持范围、迁移阶段与剩余验收；主网格已切换 |
| [M5 结构与工作区](18-m5-structure-and-workspace.md) | 结构、排序引用、筛选/冻结、基础样式、最近文件与外部冲突 |
| [M5 网格与输入](17-m5-grid-and-input.md) | 网格适配、虚拟化测量、键盘/选区/剪贴板、IME 验收边界 |
| [M4 回收与部署](16-m4-recovery-and-deployment.md) | 宿主崩溃回收、部署证据与未关闭项 |
| [M4 包内模块](12-m4-modules.md) | API v2、逻辑 ID 导入、同步初始化、版本兼容与模块验证 |
| [M4 actions 操作脚本](13-m4-actions.md) | 只读快照、二次校验、原子提交、一次撤销、桌面运行与验证 |
| [M4 独立脚本进程](11-m4-worker.md) | worker 首切片的取消、故障和资源边界；保留历史发布证据 |
| [产品范围](01-product-and-scope.md) | 长期 Beta 的产品边界和暂缓功能 |
| [系统架构](02-system-architecture.md) | 当前模块依赖及后续数据/计算设计 |
| [文件格式与兼容性](03-file-format-and-compatibility.md) | XLSQ v2、v1 历史设计、类型与保存协议；当前支持子集另见 09/25 |
| [脚本与运行时](04-scripting-and-runtime.md) | worker/模块/actions 的当前入口与历史设计、严格策略及权限边界 |
| [质量性能与交付](06-quality-performance-and-delivery.md) | 已有证据、候选包验收和未来质量门禁 |
| [架构决策](07-architecture-decisions.md) | 已选方向、取舍及待验证事项 |
| [基线与实施记录](08-baseline-and-progress.md) | 初始 102 例基线和 0.2 新验证分别记载 |
| [M3 跨表与增量计算](10-m3-calculation.md) | 0.4 语义矩阵、脚本桥、结构命令、导出与验收 |
| [M2 文件与类型](09-m2-files-and-types.md) | 0.3 支持矩阵、实现、验证与下一阶段 |

用户操作见[试用说明](../README.md)，公开限制和运行时要求见[发布说明](../RELEASE-NOTES.md)。开发时先读系统架构与实施计划，再进入专题。

历史修复版 0.2.1：关闭确认延后到 WPF Closing 事件返回后执行，并新增真实窗口关闭回归。0.2.0 的发布记录保留在实施记录第 6 节，关闭缺陷与 0.2.1 验证见第 7 节。

## 状态约定

- **当前实现**：已在源码中接入的行为；验证状态以实施记录为准。
- **目标设计**：尚待实施的格式、worker、完整桌面和 Beta 能力。
- **待验证决策**：有明确实验或里程碑，未完成取证前不视为通过。
- **后续范围**：本次 0.2 交付不包含的功能，不因已有基础类或局部测试而关闭整个任务。

历史 M1：DOM-001/002、APP-001/002/003、CMD-001/002、UI-001 的实现与交付验收完成。当时 Release 构建、Core 221 例、native/WPF smoke 和 ZIP 解压后复验通过，见实施记录第 6 节。这些是 0.2 的证据，不代表本次文档同步重新测试或当前候选包已通过全部验收。

## 0.2 历史实现

| 项目 | 实现与边界 |
| --- | --- |
| 平台与发布 | Windows x64、WPF、.NET 10；framework-dependent，需要 .NET 10 Windows Desktop Runtime x64 |
| 领域与会话 | 类型值、稀疏不可变快照、稳定会话内 SheetId、单写入门禁、单调 Revision |
| 编辑与历史 | 单格/公式栏/粘贴/清空统一提交，新增表可撤销，默认 32 MiB 历史预算 |
| 计算 | 同表依赖和错误诊断；结果须匹配会话/输入修订/计算代次/脚本代次，跨表及依赖保持 Blocked 缓存 |
| 脚本 | QuickJS-NG 0.17.0，每轮新 runtime，显式启用；任何脚本编辑撤销授权；无 worker/OS 隔离 |
| 保存 | 不可变快照、串行写入和原子替换；待编辑输入先提交，重算失败时允许保存输入与旧缓存并提示 |
| 格式 | XLSQ v1 三条目；literal 保存为文本，重开保持 Text；XLSX 导出全部公式的当前值 |
| 兼容元数据 | 保留原型受支持的样式、尺寸、隐藏、合并与图片；界面呈现与完整类型保真尚不完整 |
| 验证与产物 | 本轮 build/Core 日志位于 `artifacts/validation/release-0.2`；最终包 smoke、ZIP/hash 记录见实施记录 |

源码入口：[Desktop 目录约定](../README.md#源码目录)、[桌面组合根](../Composition/DesktopComposition.cs)、[编辑控制器](../Editing/EditorController.cs)、[文档 facade](../../QuickExcel.Core/WorkbookDocument.cs)、[会话](../../QuickExcel.Core/DocumentSession.cs)、`存储适配`（历史代码，已移除）、`包封装`（历史代码，已移除）、[原生桥](../../QuickExcel.Scripting/QuickJsSession.cs)。

## 文档维护

任务状态在实施计划统一更新，契约与验证证据留在对应专题；入口文档同步摘要和链接，历史记录标明当时范围。用户确认关闭的缺陷须注明确认日期和来源，不冒充新增自动测试或扩展其他门禁。修改格式、授权或编译边界时同步 ADR 与专题。发布包元数据记录实际 Git HEAD 和工作树状态，文档更新不能重写旧包身份或将其视为新包。
