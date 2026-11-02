# QuickExcel 0.6 交互预览版

Avalonia 迁移：Windows A0–A5 工程切换完成，A7 已移除 WPF 生产入口。Linux 有受测 Ubuntu 候选包；macOS worker 未实现，脚本保持关闭。实机验证按用户要求延期，详见 [平台状态](QuickExcel.Desktop/docs/67-avalonia-platforms.md) 与 [收尾记录](QuickExcel.Desktop/docs/68-avalonia-retirement.md)。

2026-10-30 进度同步：M1–M3 受支持子集已交付，当前处于 **M5 桌面交互收敛**；M4 严格策略/部署与 M6 恢复/发布验收仍未完成。当前状态与后续顺序统一见 [实施计划](QuickExcel.Desktop/docs/05-implementation-plan.md)。

**当前重点工程：XLSX 兼容性与 Excel 常用功能对齐。** 按用户要求优先补常见文件保真、日常编辑和公式，再推进 Table/验证/条件格式、报表与高级分析。分阶段清单见 [专项计划](QuickExcel.Desktop/docs/26-excel-compatibility-roadmap.md)；E0–E4 已完成声明范围（保留各项明确限制），E3 [筛选/排序持久化](QuickExcel.Desktop/docs/50-filter-and-sort-persistence.md) 已完成 37 个真实 Excel 案例、333 项往返及相关 WPF 检查，[Table 基础](QuickExcel.Desktop/docs/51-excel-table-basics.md) 已完成 60 种内置样式、60,336 项属性往返和真实 Excel 保存重开；[结构化引用与表编辑](QuickExcel.Desktop/docs/52-structured-references-and-table-editing.md) 已完成十阶段 Excel 对照、20,891 项往返及 170 项相关测试；[数据验证](QuickExcel.Desktop/docs/53-data-validation.md) 已完成两种日期系统共 1,370 个真实 Excel 案例、45,210 项往返和 180 项相关测试，[条件格式基础](QuickExcel.Desktop/docs/54-conditional-formatting.md) 已完成 131,982 项模型/往返、29,489 项 Excel 输出复验、205 项相关及 17/153 项 WPF 检查，[常用清洗命令](QuickExcel.Desktop/docs/55-data-cleaning.md) 已完成两种日期系统 54,786 项往返、18,432 项 Excel 输出复验及 19/153 项 WPF 检查，[分组与分类汇总](QuickExcel.Desktop/docs/56-outline-and-subtotals.md) 已完成 131,545 项往返、37,928 项 Excel 输出复验、202 项相关及 16/153 项 WPF 检查；筛选与分组混合状态保留并报告，E4 [报表、打印与容量](QuickExcel.Desktop/docs/57-e4-reporting-and-capacity.md) 已完成五项声明范围，289 项相关测试、最终 117 项专项、81 项真实 Excel 和 23 项 WPF 通过，三份 PDF/DPI 验证及容量原始数据已归档。[状态栏与格式效率](QuickExcel.Desktop/docs/49-selection-statistics-and-format-efficiency.md) 已通过 104 条真实 Excel 小数位命令、三种可见性统计及 95 项专项。[计算和可移植导出](QuickExcel.Desktop/docs/48-calculation-and-portable-export.md) 已通过 1352 条真实 Excel 公式与 270 个参数边界对照。[文本和类型公式](QuickExcel.Desktop/docs/47-text-and-type-functions.md) 已通过 1900/1904 共 4,658 条真实 Excel 对照。[日期/数学及工作日](QuickExcel.Desktop/docs/46-date-and-math-functions.md) 已通过 1900/1904 共 3,024 条真实 Excel 对照。[查找公式](QuickExcel.Desktop/docs/45-lookup-functions.md) 已通过 604 条真实 Excel 对照。[条件统计公式](QuickExcel.Desktop/docs/44-conditional-functions.md) 已通过 1,706 条真实 Excel 对照。[工作表与名称管理](QuickExcel.Desktop/docs/43-sheet-and-name-management.md) 已通过真实 Excel 对照。[录入和公式辅助](QuickExcel.Desktop/docs/42-formula-entry-and-navigation.md) 已完成真实 Excel 对照。填充和序列的真实 Excel 对照见 [填充记录](QuickExcel.Desktop/docs/41-fill-and-series.md)。内部复制/剪切、选择性粘贴和真实 Excel 对照见 [剪贴板记录](QuickExcel.Desktop/docs/40-internal-clipboard.md)。

当前主工作区使用 `QuickExcel.Grid.Avalonia`，复用共享几何、编辑控制器与 Skia 渲染；WPF 作为归档回退基线保留。此前 WPF 网格替换时，旧 DataGrid 与试验窗口已移除；编辑、筛选排序、冻结、尺寸拖动/双击自动调整、图片与合并均已接入。**横向拖动导致整桌面停顿已于 2026-10-30 经用户确认解决并标记完成。** 默认 Avalonia 自动选择渲染方式，`Start-QuickExcel-Software.cmd` 仅为本次启动强制 Windows 软件渲染，用于诊断。实现、既有专项证据与缺陷关闭记录见 [主网格迁移](QuickExcel.Desktop/docs/24-main-grid-migration.md)；真实 IME、跨屏 DPI 和完整辅助功能门禁仍待验收。

Avalonia 版本为 `0.6.0-preview.7`（旧 WPF 为 `0.6.0-preview.3`），已增加正式工作区、脚本专页、模块编辑、查找替换、Ribbon、脚本数字签名、兼容/导出入口及 XLSX 渲染修复；范围见 [正式界面](QuickExcel.Desktop/docs/19-m5-product-workspace.md)、[签名信任](QuickExcel.Desktop/docs/20-script-signatures.md)及[渲染记录](QuickExcel.Desktop/docs/25-xlsx-rendering.md)。版本号相同不代表旧发布目录包含最新源码。

默认 Windows x64 桌面表格，Linux 候选范围见平台日志；基于 Avalonia、自有类型化存储/计算引擎与 QuickJS-NG。M4 宿主崩溃后的 profile/payload 回收已实现，本机部署验证五项通过、直接长路径 EXE 启动一项失败；后续环境见 [回收与部署验证](QuickExcel.Desktop/docs/16-m4-recovery-and-deployment.md)。工作簿恢复目前仅支持手动打开副本，自动恢复后端仍属 M6。

## 试用

`Start-QuickExcel.cmd` 启动经过核验的活动 Avalonia 包；没有活动包时构建 Web 资源并从 Release 源码运行。发布命令与回退方法见 [Windows 发布记录](QuickExcel.Desktop/docs/66-avalonia-windows-release.md)。

新 Avalonia 包自带 .NET 10 与独立 worker，PDF.js/Monaco 需要 Microsoft WebView2 Evergreen Runtime；离线安装说明随包。旧 `artifacts/QuickExcel/QuickExcel.exe` 是 WPF 快照，仍需要 .NET Windows Desktop Runtime，不能视为最新版本。

- 新建及保存使用 **XLSQ v2**：数字、文本、布尔、标准错误、空文本、空白和公式缓存分别保存；保留 1900/1904 日期系统、数字格式及稳定工作表身份。
- v1 打开时迁移，不猜测恢复历史上已丢失的类型；首次保存必须选新路径，保留原文件。
- 不支持的 XLSX 部件和属性进入兼容报告。工作簿内容未变时保留原 XLSX 字节；编辑后普通保存会阻止不确定的损失，可明确选择另存受支持副本。
- XLSX 可直接保存已支持内容、公式和文档属性。冻结、每表缩放、活动表/格、滚动位置与 hidden/veryHidden 持久化；普通取消隐藏不暴露 veryHidden。冻结/缩放/可见性可撤销且产生 dirty，导航只随下一次保存记录，范围见 [视图契约](QuickExcel.Desktop/docs/33-xlsx-workbook-views.md)。
- XLSX 导出可选值快照或可移植公式；脚本与跨表传递依赖冻结为值，独立标准公式可保留。缺失的冻结结果阻止导出，旧结果须明确选择。
- 存储读取、迁移、保存不执行脚本或恢复授权；桌面打开时可在签名与证书信任验证通过后启用脚本。未签名文件保持停用，恢复副本始终不自动启用，详见[签名信任](QuickExcel.Desktop/docs/20-script-signatures.md)。文件级 API v2 支持包内 ES 模块执行；import 只接受清单声明的逻辑 ID。API v1 保留旧单脚本语义。

脚本每轮在独立 worker 执行，可随时停止。已绑定的 actions 读取只读工作簿快照，批量写值或公式；宿主校验全部变更后原子提交，可一次撤销，失败不留部分写入。独立 UDF 已按原顺序合批，消息上限为 16 MiB；Windows 公式与 actions worker 已强制使用无能力 AppContainer，回环与 DNS 拒绝已取得本机 WFP 证据。严格执行策略、部署矩阵及 M6 OS 门禁仍待完成。本版仅供可信脚本使用，不等于完整 M4 或 Beta。完整桌面交互与自动恢复属 M5/M6。见 [合批与容量](QuickExcel.Desktop/docs/14-m4-batching.md)、[隔离 ADR](QuickExcel.Desktop/docs/15-m4-os-isolation-adr.md)、[actions 契约与验收](QuickExcel.Desktop/docs/13-m4-actions.md)及[发布说明](QuickExcel.Desktop/RELEASE-NOTES.md)。

## 仓库边界

QuickExcel 与同级 `../ExcelIO` 是独立 Git 仓库。QuickExcel 已移除 ExcelIO 和 ExcelIO.Formula 依赖，构建、测试与发布无需同级库；不修改、复制或提交其源码。清理范围与兼容验证见 [依赖移除记录](QuickExcel.Desktop/docs/58-excelio-removal.md)。

| 目录 | 职责 |
| --- | --- |
| `QuickExcel.Core` | 纯 .NET 领域快照、会话、命令、历史和端口 |
| `QuickExcel.Infrastructure` | 类型化 XLSX/XLSQ v2、v1 导入迁移、计算与运行时适配 |
| `QuickExcel.Scripting` | QuickJS 托管桥和固定来源的原生引擎 |
| `QuickExcel.ScriptContracts` / `QuickExcel.ScriptWorker` | 有界 IPC 协议与独立脚本执行进程 |
| `QuickExcel.Desktop.Avalonia` / `Grid.Avalonia` | 默认桌面、原生网格、Monaco/PDF.js 与平台组合根 |
| `QuickExcel.Presentation` / `Grid.Shared` | 无 UI 的共享编辑、工作区和几何规则 |
| `QuickExcel.Desktop` | 产品/架构文档、发布历史和保留许可证；不再包含 WPF 应用项目 |
| `QuickExcel.Grid.Wpf` | 旧网格许可证与来源清单归档；生产使用 Grid.Avalonia/Grid.Shared |
| `QuickExcel.Core.Test` | 领域、会话、存储、计算和格式异常回归 |

## 构建

源码构建需要 Windows x64、.NET 10 SDK 和 Node.js/npm；发布脚本使用 PowerShell 7。首次原生构建还需 CMake、Visual Studio C++ 工具及下载固定源码的网络连接，见[来源说明](QuickExcel.Scripting/native/SOURCES.md)。

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter "FullyQualifiedName~WebDraftContractTests|FullyQualifiedName~M5InteractionTests" -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

所有托管命令均须指定 `--artifacts-path`，让引用项目的构建输出也进入本仓库。上面的测试是定向示例；日常变更只运行相关测试和 Avalonia 专项，不逐次执行全量套件。发布脚本生成独立新发布树、ZIP、SHA-256 和 release.json；不把旧 smoke 文件当成本次通过证据。无 SDK 干净机器、安装器和程序/安装包签名验收尚未完成；工作簿脚本签名是另一项已实现能力。

- [试用操作](QuickExcel.Desktop/README.md)
- [实施计划](QuickExcel.Desktop/docs/05-implementation-plan.md)
- [M2 支持矩阵与证据](QuickExcel.Desktop/docs/09-m2-files-and-types.md)

当前计算范围、脚本描述符与验收见 [M3 支持矩阵](QuickExcel.Desktop/docs/10-m3-calculation.md)。

Linux 发布使用 `scripts/publish-avalonia-unix.ps1 -Rid linux-x64 -CreateArchive`，先在目标平台构建固定来源 native。Linux worker 必须通过 Landlock/seccomp/cgroup 的逐轮核验，Secret Service 不可用时不持久信任。macOS 发布工具只准备工程预览，不能据此声明平台支持。

WPF 源码保存在标签 `avalonia-wpf-baseline-20261102`，受测归档包与校验清单在 `artifacts/avalonia/rollback`。回退用 `scripts/select-quickexcel-package.ps1 -PackageDirectory <归档目录>`；不重新编译或在当前宿主中加载 WPF。
