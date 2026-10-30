# QuickExcel 0.6 交互预览版

2026-10-30 进度同步：M1–M3 受支持子集已交付，当前处于 **M5 桌面交互收敛**；M4 严格策略/部署与 M6 恢复/发布验收仍未完成。当前状态与后续顺序统一见 [实施计划](QuickExcel.Desktop/docs/05-implementation-plan.md)。

**当前重点工程：XLSX 兼容性与 Excel 常用功能对齐。** 按用户要求优先补常见文件保真、日常编辑和公式，再推进 Table/验证/条件格式、报表与高级分析。分阶段清单见 [专项计划](QuickExcel.Desktop/docs/26-excel-compatibility-roadmap.md)；E0/E1 及 XC-0201/0202/0203 已完成声明范围，接续工作表与名称管理。[录入和公式辅助](QuickExcel.Desktop/docs/42-formula-entry-and-navigation.md) 已完成真实 Excel 对照。填充和序列的真实 Excel 对照见 [填充记录](QuickExcel.Desktop/docs/41-fill-and-series.md)。内部复制/剪切、选择性粘贴和真实 Excel 对照见 [剪贴板记录](QuickExcel.Desktop/docs/40-internal-clipboard.md)。

主工作区已直接使用 `QuickExcel.Grid.Wpf`，旧 DataGrid 与试验窗口已移除；编辑、筛选排序、冻结、尺寸拖动/双击自动调整、图片与合并均已接入。**横向拖动导致整桌面停顿已于 2026-10-30 经用户确认解决并标记完成。** 默认 WPF 自动渲染，软件渲染仅作诊断。实现、既有专项证据与缺陷关闭记录见 [主网格迁移](QuickExcel.Desktop/docs/24-main-grid-migration.md)；真实 IME、跨屏 DPI 和完整辅助功能门禁仍待验收。

源码版本仍为 `0.6.0-preview.3`，已增加正式工作区、脚本专页、模块编辑、查找替换、Ribbon、脚本数字签名、兼容/导出入口及 XLSX 渲染修复；范围见 [正式界面](QuickExcel.Desktop/docs/19-m5-product-workspace.md)、[签名信任](QuickExcel.Desktop/docs/20-script-signatures.md)及[渲染记录](QuickExcel.Desktop/docs/25-xlsx-rendering.md)。版本号相同不代表旧发布目录包含最新源码。

Windows x64 桌面表格，基于 WPF、同级 ExcelIO 源码库与 QuickJS-NG。M4 宿主崩溃后的 profile/payload 回收已实现，本机部署验证五项通过、直接长路径 EXE 启动一项失败；后续环境见 [回收与部署验证](QuickExcel.Desktop/docs/16-m4-recovery-and-deployment.md)。工作簿恢复目前仅支持手动打开副本，自动恢复后端仍属 M6。

## 试用

测试当前源码时，按下方命令构建 Release 后运行 `artifacts/build/bin/QuickExcel.Desktop/release/QuickExcel.exe`。`Start-QuickExcel.cmd` 当前从源码执行 `dotnet run`，未指定配置时为 Debug，需要 SDK；性能验收统一使用 Release。

已有 `artifacts/QuickExcel/QuickExcel.exe` 是独立发布快照，不会随源码构建自动更新，须核对同目录 `BUILD-METADATA.json`。发布包需要 **.NET 10 Windows Desktop Runtime x64**，采用 framework-dependent 发布；本次文档同步不生成新包。

- 新建及保存使用 **XLSQ v2**：数字、文本、布尔、标准错误、空文本、空白和公式缓存分别保存；保留 1900/1904 日期系统、数字格式及稳定工作表身份。
- v1 打开时迁移，不猜测恢复历史上已丢失的类型；首次保存必须选新路径，保留原文件。
- 不支持的 XLSX 部件和属性进入兼容报告。工作簿内容未变时保留原 XLSX 字节；编辑后普通保存会阻止不确定的损失，可明确选择另存受支持副本。
- XLSX 可直接保存已支持内容、公式和文档属性。冻结、每表缩放、活动表/格、滚动位置与 hidden/veryHidden 持久化；普通取消隐藏不暴露 veryHidden。冻结/缩放/可见性可撤销且产生 dirty，导航只随下一次保存记录，范围见 [视图契约](QuickExcel.Desktop/docs/33-xlsx-workbook-views.md)。
- XLSX 导出可选值快照或可移植公式；脚本与跨表传递依赖冻结为值，独立标准公式可保留。缺失的冻结结果阻止导出，旧结果须明确选择。
- 存储读取、迁移、保存不执行脚本或恢复授权；桌面打开时可在签名与证书信任验证通过后启用脚本。未签名文件保持停用，恢复副本始终不自动启用，详见[签名信任](QuickExcel.Desktop/docs/20-script-signatures.md)。文件级 API v2 支持包内 ES 模块执行；import 只接受清单声明的逻辑 ID。API v1 保留旧单脚本语义。

脚本每轮在独立 worker 执行，可随时停止。已绑定的 actions 读取只读工作簿快照，批量写值或公式；宿主校验全部变更后原子提交，可一次撤销，失败不留部分写入。独立 UDF 已按原顺序合批，消息上限为 16 MiB；公式与 actions worker 已强制使用无能力 AppContainer，回环与 DNS 拒绝已取得本机 WFP 证据。严格执行策略、部署矩阵及 M6 OS 门禁仍待完成。本版仅供可信脚本使用，不等于完整 M4 或 Beta。完整桌面交互与自动恢复属 M5/M6。见 [合批与容量](QuickExcel.Desktop/docs/14-m4-batching.md)、[隔离 ADR](QuickExcel.Desktop/docs/15-m4-os-isolation-adr.md)、[actions 契约与验收](QuickExcel.Desktop/docs/13-m4-actions.md)及[发布说明](QuickExcel.Desktop/RELEASE-NOTES.md)。

## 仓库边界

QuickExcel 与同级 `../ExcelIO` 是独立 Git 仓库。ExcelIO 是源码构建依赖；普通 QuickExcel 任务不修改、复制或提交同级库的源码。

| 目录 | 职责 |
| --- | --- |
| `QuickExcel.Core` | 纯 .NET 领域快照、会话、命令、历史和端口 |
| `QuickExcel.Infrastructure` | 类型化 XLSX/XLSQ v2、保留的 v1 适配、计算与运行时适配 |
| `QuickExcel.Scripting` | QuickJS 托管桥和固定来源的原生引擎 |
| `QuickExcel.ScriptContracts` / `QuickExcel.ScriptWorker` | 有界 IPC 协议与独立脚本执行进程 |
| `QuickExcel.Desktop` | WPF、组合根、编辑控制器和集成 smoke |
| `QuickExcel.Grid.Wpf` | 固定上游来源的通用主网格控件、内容/选区分层绘制与坐标映射；不引用领域或脚本 |
| `QuickExcel.Core.Test` | 领域、会话、存储、计算和格式异常回归 |

## 构建

需要 Windows x64、.NET 10 SDK 和同级 ExcelIO。首次原生构建还需 CMake、Visual Studio C++ 工具及下载固定源码的网络连接，见[来源说明](QuickExcel.Scripting/native/SOURCES.md)。

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

所有托管命令均须指定 `--artifacts-path`，让引用项目的构建输出也进入本仓库。以上完整验证用于基线/发布验收；日常变更只运行相关测试和 WPF 专项，不逐次执行全量套件。发布脚本生成独立新发布树、ZIP、SHA-256 和 BUILD-METADATA；不把旧 smoke 文件当成本次通过证据。无 SDK 干净机器、安装器和程序/安装包签名验收尚未完成；工作簿脚本签名是另一项已实现能力。

- [试用操作](QuickExcel.Desktop/README.md)
- [实施计划](QuickExcel.Desktop/docs/05-implementation-plan.md)
- [M2 支持矩阵与证据](QuickExcel.Desktop/docs/09-m2-files-and-types.md)

当前计算范围、脚本描述符与验收见 [M3 支持矩阵](QuickExcel.Desktop/docs/10-m3-calculation.md)。
