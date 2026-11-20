# QuickExcel 5.1.0

QuickExcel 是基于 .NET 10、Avalonia 和 Skia 的桌面电子表格应用，提供自有的类型化 XLSX/XLSQ 存储、增量公式计算，以及在独立沙箱中运行的 JavaScript / TypeScript 脚本。默认交付平台为 Windows x64；Linux x64 与 macOS arm64 的支持范围单独验证。

当前源码版本以 [桌面项目的 Version](QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj) 为准。4.0.0 的 Windows 发布采用 **Native AOT 脚本 worker + 部分裁剪的自包含桌面运行时**，JSC、ICU 与 TypeScript 仅随 worker 分发。实现与交付证据见 [4.0 发布记录](QuickExcel.Desktop/docs/84-native-aot-and-trimming.md)。

## 最近更新

以下版本均更新于 2026-11-20，按版本从新到旧排列：

| 版本 | 更新内容 |
| --- | --- |
| **5.1.0** | 删除行列不再弹出确认；选区删除支持左移、上移、整行和整列，支持撤销。新增列表式名称管理器、作用域筛选、批注和右键创建命名区域，支持 XLSX/XLSQ 保存及脚本访问。见[删除单元格与名称管理器](QuickExcel.Desktop/docs/87-delete-cells-and-name-manager.md)。 |
| **5.0.0** | 统一 XLSX、CSV/TSV、ODBC 和 SQLite 数据导入导出，支持工作表选择、重名处理、当前工作簿导入与撤销、数据库事务写入；兼容报告聚合同类问题并保留更多 Excel 元数据。新增 Windows/macOS GitHub Release 自动打包流程。见[数据交换](QuickExcel.Desktop/docs/86-data-import-export.md)、[兼容性记录](QuickExcel.Desktop/docs/39-xlsx-diagnostic-closure.md)和[发布流程](QuickExcel.Desktop/docs/85-github-actions-release.md)。 |
| **4.2.2** | 另存受支持副本后，兼容状态与实际保存的文件同步；后续保存及撤销/重做不再沿用原文件的过期兼容信息。 |
| **4.2.1** | 修复 ClickOnce 更新后的文件关联与桌面快捷方式；新增 XLSX、CSV 专用文件图标。 |
| **4.2.0** | 只选表头建立筛选时，向下纳入所选列中的连续记录；保留明确多行选区的范围，阻止对仅含表头的区域排序。 |
| **4.1.0** | ClickOnce 安装包改用 JSC 完整对应源码下载地址，保留源码获取、修改说明和构建信息；普通发布归档仍保留源码包。 |
| **4.0.0** | Windows 脚本 worker 改为 Native AOT，桌面使用部分裁剪的自包含运行时，消除桌面根目录重复的 JSC/ICU/TypeScript 载荷。 |

5.1.0 已通过相关 Core/存储测试、7 项发布 AOT worker 测试、220 项界面检查及 36 项生产自检；验证范围和证据见[专题记录](QuickExcel.Desktop/docs/87-delete-cells-and-name-manager.md#定向验证)。

## 快速开始

在 Windows 上从仓库根目录运行：

```powershell
.\Start-QuickExcel.cmd
# 以禁用脚本模式启动，仍可查看和编辑工作簿及脚本源码：
.\Start-QuickExcel.cmd --no-scripts
```

启动器读取 `artifacts/avalonia/active-windows.json`，核对发布清单与入口文件身份后启动已验证的活动包；没有活动包时构建 Web 资源并从 Release 源码运行，因此需要下文的开发工具链。当前 Windows 程序入口为 `quickxl.exe`。显示问题可通过启动环境变量 `QUICKEXCEL_SOFTWARE_RENDERING=1` 启用 Windows 软件渲染进行诊断。

Windows 发布包自带 .NET 10，运行无需安装 SDK。Monaco 脚本编辑器与 PDF.js 预览需要 WebView2 Evergreen Runtime；ClickOnce 的 `setup.exe` 会检测该组件，缺少时使用随包的在线引导程序下载并安装，届时需要联网。安装、签名及更新方法见 [Windows 发布与 ClickOnce](QuickExcel.Desktop/docs/66-avalonia-windows-release.md)。

开始页或「文件 → 设置」提供文件关联、历史记录、计算、工作表标签位置、脚本禁用选项、脚本 API 浏览器与第三方许可证。

## GitHub Release 自动打包

推送与桌面项目 `Version` 一致的 `vMAJOR.MINOR.PATCH` 标签（例如 `v5.1.0`），[Release Windows and macOS](.github/workflows/release.yml) 会分别构建 Windows x64 ZIP 和 macOS arm64 tar.gz，运行生产包自检与相关 worker 测试，并在两端都成功后自动发布到 GitHub Release，附带 SHA-256 校验文件。也可从 Actions 手动输入已有标签重跑。Linux 暂不参与此发布流程。

macOS 当前交付原始目录压缩包，尚不是 Developer ID 签名、公证的 `.app`。触发方式、产物与失败恢复见 [GitHub Actions 发布说明](QuickExcel.Desktop/docs/85-github-actions-release.md)。

## 主要功能

以下均指已实现的支持子集，具体参数、容量及文件往返边界以专题文档为准。

| 领域 | 当前能力 |
| --- | --- |
| 文件与兼容 | XLSQ v2、XLSX 读写及 Strict OOXML 支持子集；CSV/TSV 导入预览、编码与逐列类型选择、导出；兼容报告、外部文件变更检测和原子保存 |
| 数据导入导出 | [统一视图](QuickExcel.Desktop/docs/86-data-import-export.md)汇集 XLSX、CSV/TSV、ODBC 数据库和 SQLite 文件；工作表重名预览、当前工作簿导入与撤销、数据库新建/追加及事务写入 |
| 日常编辑 | 撤销/重做、查找替换、内部复制/剪切、选择性与运算粘贴、可见行复制粘贴、填充/序列、定位条件、直接删除行列、选区左移/上移删除、工作表管理 |
| 名称管理 | 名称列表、数值预览、搜索与作用域筛选、新建/编辑/删除、批注、右键选区创建命名区域、XLSX/XLSQ 保存及脚本访问 |
| 格式与内容 | 数字格式、样式、合并、冻结/缩放、富文本分段编辑、超链接、传统批注与线程评论、图片和绘图对象 |
| 数据处理 | 多键排序、颜色/前后 N 项/日期分组筛选、Excel Table 与结构化引用、数据验证、条件格式、分列、去重、快速填充、分组与分类汇总 |
| 公式与分析 | 跨表引用、名称、稀疏依赖与增量计算；条件统计、查找、日期/数学/文本函数、LET、动态数组、基础透视表；公式引用/从属追踪、错误定位与逐步求值 |
| 报表与打印 | 页面设置、分页预览、图表、Skia 共享渲染、PDF 导出及原生打印入口 |
| 脚本与恢复 | Monaco 编辑、TypeScript 检查、UDF、Actions、Events、初始化入口、调试控制台、脚本数字签名、会话存储与虚拟文件、自动恢复草稿 |

E0–E5 及 E6/E7 已完成各自声明的交付子集；E8/E9 的分析、复杂模板和容量扩展仍在计划中。功能清单与后续范围见 [Excel 对齐路线图](QuickExcel.Desktop/docs/26-excel-compatibility-roadmap.md)、[E5 实施记录](QuickExcel.Desktop/docs/74-e5-implementation-log.md)和 [常用功能计划](QuickExcel.Desktop/docs/82-excel-common-features-plan.md)。

## 文件、保存与恢复

- **XLSQ v2** 是新建工作簿和脚本工作簿的主格式，分别保存数字、文本、布尔、错误、空白、空文本与公式缓存，并保留日期系统和工作表身份。旧 XLSQ v1 数据导入时迁移，首次保存须使用新路径，不猜测恢复历史上已丢失的类型。
- **XLSX** 可直接保存已支持内容、公式、视图和元数据。不支持的部件或语义进入兼容报告；内容未变时保留原包字节，编辑后若存在不确定损失则阻止普通保存，可明确选择另存受支持副本。未知公式和旧缓存不会被冒充为已支持计算或新鲜结果。
- **XLSX 导出** 可选值快照或可移植公式。脚本及其依赖按导出策略冻结为值；缺失结果阻止导出，使用旧结果需要明确选择。脚本源码及持久脚本存储使用 XLSQ 保存。
- **自动恢复** 每 30 秒尝试记录有未保存编辑的工作簿，保留两代完整草稿。启动流程和文件菜单提供恢复列表；恢复为脱离原路径的新会话，保持未保存且脚本停用。草稿包含已提交到文档会话的内容，不包含编辑框尚未提交的输入、`sessionStorage` 或临时虚拟文件。完整契约与故障验证见 [E6/E7 实施记录](QuickExcel.Desktop/docs/83-e6-e7-implementation-log.md)。

当前模型上限为每表 **200,000 行 × 2,048 列**、每簿 **64 张表、1,000,000 个实际存储单元格**；各操作另有内存、历史与时间预算，这些数字不等于全容量性能验收。文件支持与计算边界分别见 [文件和类型](QuickExcel.Desktop/docs/09-m2-files-and-types.md)、[计算契约](QuickExcel.Desktop/docs/10-m3-calculation.md)及后续专题。

## 脚本模型

生产脚本引擎为 JavaScriptCore；TypeScript 5.9.3 在独立隔离编译 worker 中进行完整检查和编译。脚本不能直接访问宿主 CLR、操作系统文件或网络，模块导入只解析工作簿清单声明的逻辑 ID。

新工作簿使用 **文件 API v4、IPC 11**：一个可选初始化主入口，UDF、Actions、Events 各有独立角色，并从模块导出自动发现。v4 不使用 `host.functions.register` 或 `host.events.on`。现有 API v3 文件保留版本隔离的兼容执行；文件 API v1/v2 仅供源码检查，须显式改写为 v3 才能执行，保存前需完成改写或移除全部脚本，不自动迁移源码。文件 API 版本与函数描述符的参数语义版本相互独立。

Actions/Events 通过 `QuickExcel.Workbook` 等对象访问宿主拥有的事务，可以读取本事务先前的写入；工作簿和虚拟文件变更统一校验后原子提交，失败全部回滚，可撤销的变更合为一次历史记录。`localStorage` 随 XLSQ 保存，`sessionStorage` 随文档会话释放；两者的事务及历史规则见 [脚本存储](QuickExcel.Desktop/docs/81-script-storage.md)。`//xlsq/`、`//temp/`、`//mem/` 是逻辑虚拟文件空间，不是本机路径。

在网格中右键选区并选择「创建命名区域…」，即可在脚本中通过名称访问。例如工作簿作用域的 `ELN` 使用 `workbook.getNamedItem("ELN").getRange()`；工作表作用域名称使用 `workbook.getWorksheet("配额表").getNamedItem("ELN").getRange()`。名称重命名后已有脚本句柄仍有效，删除后重新创建同名区域不会复用旧句柄身份。

Windows/macOS 在每个已授权工作簿会话中分别复用 UDF 与 Actions/Events 的隔离 worker 和上下文；UDF 不共享初始化状态。失败或撤权后丢弃上下文，不自动重放初始化。控制台、编译器及 Linux worker 保持一次性执行；Linux 当前拒绝持久初始化绑定。

存储读取本身不执行脚本。脚本启用受授权、签名信任与执行策略约束；设置中的「禁用脚本」、`--no-scripts` 或 `QUICKEXCEL_NO_SCRIPTS=1` 均可禁止执行和编译，优先于已信任签名。API 浏览器展示当前能力与成员签名，未实现成员或参数组合显式拒绝。

入口示例见 [初始化、自动发现与编辑器](QuickExcel.Desktop/docs/80-script-entrypoints-and-editor.md)，对象支持矩阵见 [脚本 API](QuickExcel.Desktop/docs/79-script-api-redesign.md)，上下文与事务计算见 [UDF 隔离](QuickExcel.Desktop/docs/77-udf-context-isolation-and-transaction-calculation.md)。

## 构建与验证

Windows 开发需要 .NET 10 SDK、Node.js/npm、PowerShell 7，以及原生引擎所需的 C++ 工具链。**Windows 正式发布要求 .NET SDK 10.0.404 或更新版本**，Native AOT 需要 Visual Studio C++ 工具和 Windows SDK；完整原生依赖、固定源码与哈希见 [JSC 构建说明](tools/QuickExcel.JscProbe/README.md)及 [来源记录](QuickExcel.Scripting/native/jsc/SOURCES.md)。首次还原或构建需下载依赖。

从仓库根目录执行：

```powershell
.\scripts\build-avalonia-web.ps1
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Desktop.Avalonia\QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path .\artifacts\build
```

按修改范围选用定向检查；下面是工作区交互和原生桥检查示例，不要求每次全部运行：

```powershell
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter "FullyQualifiedName~WebDraftContractTests|FullyQualifiedName~M5InteractionTests" -v minimal
.\scripts\test-avalonia-workspace.ps1 -Mode Workspace
# 网格变更使用 -Mode Grid；性能基准单独运行：
.\scripts\test-avalonia-grid-benchmark.ps1
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
```

所有托管构建、测试、运行和发布命令都须指定 `--artifacts-path`，将输出保留在本仓库 `artifacts` 下。日常功能或修复只运行相关测试，不逐次执行全量回归；性能采样与功能测试分开运行。Windows AOT worker 的集成验证须通过 `QuickExcelPublishedWorkerDirectory` 使用真实发布载荷，普通开发构建不能替代该验证。

相关检查通过后，功能或修复交付必须更新本地启动包：

```powershell
.\scripts\build-quickexcel.ps1 -SkipTests -RunSmoke
# 如需同时生成 ZIP 和 SHA-256：
# .\scripts\build-quickexcel.ps1 -SkipTests -RunSmoke -CreateArchive
```

发布脚本分别发布 AOT worker 和自包含桌面，在全新目录执行生产自检，成功后才更新 `artifacts/avalonia/active-windows.json`。失败保留之前的活动包；普通源码构建不会更新启动包。`-OutputDirectory` 必须位于本仓库 `artifacts` 下。ClickOnce 使用 `scripts/build-clickonce.ps1 -CertificateThumbprint <thumbprint>`，签名证书准备、构建前置条件和分发验证见 [Windows 发布文档](QuickExcel.Desktop/docs/66-avalonia-windows-release.md)。

## 平台与验收状态

| 平台 | 当前工程范围 | 仍需独立验收的事项 |
| --- | --- | --- |
| Windows x64 | 默认 Avalonia 包；Native AOT worker、零能力 AppContainer、签名与 ClickOnce 分发；4.0 包及其启动器/ClickOnce 载荷已通过生产自检 | 干净设备安装、完整部署与最终 OS 安全门禁、物理输入/打印等 |
| Linux x64 | 受测 Ubuntu 候选包；Landlock/seccomp/systemd 用户 scope 隔离、Secret Service 信任、CUPS 打印；托管 worker | 一次性脚本上下文不支持持久初始化；VM 中文输入候选窗问题及实机验收仍未关闭 |
| macOS arm64 | 原生 JSC、默认拒绝的 Seatbelt 隔离、持久 UDF/Actions worker、本机发布自检；托管 worker | Developer ID 签名/公证、签名包持久信任和实机验收仍未完成 |

Linux/macOS 通过 `scripts/publish-avalonia-unix.ps1` 发布，分别传入 `-Rid linux-x64` 或 `-Rid osx-arm64`，可加 `-CreateArchive` 归档。需先在目标平台构建原生引擎与 worker 沙箱；具体命令和前置条件见 [AGENTS.md](AGENTS.md)与 [平台实施记录](QuickExcel.Desktop/docs/67-avalonia-platforms.md)。macOS 的 256 MiB 进程限制由 footprint 轮询看门狗实施，允许有界超调，不是 cgroup 等价硬限。平台密钥保护不可用时仅允许明确的一次性授权。

M4/M5/M6 的完整验收尚未关闭，真实 IME、跨屏 DPI、辅助功能、物理打印和干净设备验证仍按既有安排延期。`JSC-MEM-001`（64 MiB 总引擎分配硬限额）保持开放、按用户决定延期；进程限额或 GC 阈值不代表它已实现。自动恢复已交付的子集也不等于全容量、断电设备和全部故障矩阵通过。

## 仓库结构

QuickExcel 与同级 `../ExcelIO` 是独立 Git 仓库。当前项目不依赖 ExcelIO 或 ExcelIO.Formula，构建、测试和发布无需同级库；相关变更留在各自仓库，见 [依赖移除记录](QuickExcel.Desktop/docs/58-excelio-removal.md)。

| 目录 | 职责 |
| --- | --- |
| [QuickExcel.Core](QuickExcel.Core/README.md) | 纯 .NET 领域模型、文档会话、命令、历史、公式与端口 |
| `QuickExcel.Infrastructure` | XLSX/XLSQ 存储、迁移、计算、恢复及隔离 worker 适配 |
| `QuickExcel.Scripting` | JSC 托管桥、TypeScript 编译支持、原生源码与运行时冒烟检查 |
| `QuickExcel.ScriptContracts` / `QuickExcel.ScriptWorker` | BCL-only IPC 契约与独立脚本执行进程 |
| `QuickExcel.Desktop.Avalonia` / `QuickExcel.Grid.Avalonia` | 默认桌面宿主、原生网格与平台组合根 |
| `QuickExcel.Presentation` / `QuickExcel.Grid.Shared` | 无 UI 的编辑控制器、工作区规则与稀疏几何算法 |
| `QuickExcel.Rendering` / `QuickExcel.Rendering.Skia` | 共享渲染契约、Skia 绘制与 PDF 输出 |
| `QuickExcel.Web` | 本地 Monaco / PDF.js Web 资源与构建入口 |
| `QuickExcel.Core.Test` / `QuickExcel.Rendering.Test` / `QuickExcel.Rendering.Smoke` | 领域、存储、计算、脚本、恢复与渲染定向验证 |
| `tools/QuickExcel.ApiGen` | 从固定 Office Scripts 定义生成脚本声明、API 目录与能力清单；生成文件不手工编辑 |
| `QuickExcel.Desktop` | 产品/架构文档、历史发布记录与许可证 |
| `scripts` / `tools` / `third_party` | 构建发布、专项验证、开发工具与保留来源的第三方组件 |

`artifacts`、构建缓存、原生下载源码及产物不进入 Git。WorkerFixture、SandboxHostFixture、RecoveryFixture 和平台探针等开发/故障注入工具不得进入正式包。

WPF 生产项目已移除，旧源码保留在 `avalonia-wpf-baseline-20261102` 标签，许可证与来源归档保留在 `QuickExcel.Grid.Wpf` 等目录。回退使用已验证的完整归档包和 `scripts/select-quickexcel-package.ps1`，不恢复旧源码构建；详见 [WPF 退役记录](QuickExcel.Desktop/docs/68-avalonia-retirement.md)。

## 文档与第三方声明

- [文档索引](QuickExcel.Desktop/docs/README.md)：架构、格式、交互与实现专题；带日期的旧记录保留当时状态。
- [E6/E7 实施记录](QuickExcel.Desktop/docs/83-e6-e7-implementation-log.md)：数据交换、公式审核、内容编辑与自动恢复。
- [脚本入口与编辑器](QuickExcel.Desktop/docs/80-script-entrypoints-and-editor.md)、[对象 API 支持矩阵](QuickExcel.Desktop/docs/79-script-api-redesign.md)：当前脚本契约与限制。
- [Windows 发布](QuickExcel.Desktop/docs/66-avalonia-windows-release.md)、[4.0 AOT 与裁剪](QuickExcel.Desktop/docs/84-native-aot-and-trimming.md)、[平台状态](QuickExcel.Desktop/docs/67-avalonia-platforms.md)：包身份、验证证据与未完成项。
- [渲染来源](QuickExcel.Rendering.Skia/SOURCES.md)、[JSC 来源](QuickExcel.Scripting/native/jsc/SOURCES.md)、[Office Scripts 定义来源](third_party/office-scripts-docs/README.md)：固定版本、哈希、许可证与修改说明。

第三方版权与许可证可在设置页及发布包的许可证文件中查看。引入组件时须同步更新设置页、随包 notices 和完整许可证，保留 FastWpfGrid、Office Scripts 定义等上游署名与修改说明。

## 相关开源仓库

下表汇总主要第三方组件、定义来源及构建/测试依赖，历史组件在用途栏注明。固定版本、许可证和本地修改以 [第三方声明](THIRD-PARTY-LICENSES.md)、各组件的来源记录及随包许可证为准。

| 项目 / 组件 | 在 QuickExcel 中的用途 | 开源仓库 |
| --- | --- | --- |
| .NET Runtime | 桌面运行时、基础类库与 Native AOT worker 运行时支持 | [dotnet/runtime](https://github.com/dotnet/runtime) |
| Avalonia | 跨平台桌面界面、控件与主题 | [AvaloniaUI/Avalonia](https://github.com/AvaloniaUI/Avalonia) |
| Avalonia.Controls.WebView | 嵌入 Monaco 编辑器与 PDF.js 预览的原生 WebView 控件 | [AvaloniaUI/Avalonia.Controls.WebView](https://github.com/AvaloniaUI/Avalonia.Controls.WebView) |
| Avalonia.BuildServices | Avalonia 构建期依赖 | [AvaloniaUI/Avalonia.BuildServices](https://github.com/AvaloniaUI/Avalonia.BuildServices) |
| MicroCom | Avalonia 原生 COM 互操作依赖 | [kekekeks/MicroCom](https://github.com/kekekeks/MicroCom) |
| Tmds.DBus | Linux D-Bus 平台集成依赖 | [tmds/Tmds.DBus](https://github.com/tmds/Tmds.DBus) |
| ANGLE（Avalonia 分支） | Windows 图形后端 | [AvaloniaUI/angle](https://github.com/AvaloniaUI/angle) |
| SkiaSharp / HarfBuzzSharp | 图形与文字塑形的 .NET 绑定 | [mono/SkiaSharp](https://github.com/mono/SkiaSharp) |
| Skia | 网格、图片与 PDF 输出使用的底层绘图引擎 | [google/skia](https://github.com/google/skia) |
| HarfBuzz | 底层文字塑形引擎 | [harfbuzz/harfbuzz](https://github.com/harfbuzz/harfbuzz) |
| RichTextKit | 经本地适配的富文本布局、换行与文字塑形集成 | [toptensoftware/RichTextKit](https://github.com/toptensoftware/RichTextKit) |
| FastWpfGrid | 共享网格几何与算法的来源，保留上游署名 | [janproch/fastwpfgrid](https://github.com/janproch/fastwpfgrid) |
| LibTiff.NET | TIFF 图片读取 | [BitMiracle/libtiff.net](https://github.com/BitMiracle/libtiff.net) |
| PdfPig | PDF 读取与内容检查 | [UglyToad/PdfPig](https://github.com/UglyToad/PdfPig) |
| Markdown.Avalonia | Tight、ColorDocument、ColorTextBlock 组件，用于许可证页原生 Markdown 渲染 | [whistyun/Markdown.Avalonia](https://github.com/whistyun/Markdown.Avalonia) |
| Fluent UI System Icons | Ribbon 等界面的图标来源 | [microsoft/fluentui-system-icons](https://github.com/microsoft/fluentui-system-icons) |
| AngleSharp | HTML 内容解析 | [AngleSharp/AngleSharp](https://github.com/AngleSharp/AngleSharp) |
| OpenMcdf | 复合文档格式读取 | [openmcdf/openmcdf](https://github.com/openmcdf/openmcdf) |
| JavaScriptCore / WebKit | 独立隔离 worker 中的 JavaScript 引擎 | [WebKit/WebKit](https://github.com/WebKit/WebKit) |
| ICU | JSC 使用的 Unicode 与国际化支持；macOS 链接系统 ICU | [unicode-org/icu](https://github.com/unicode-org/icu) |
| mimalloc | 原生引擎构建中的内存分配器 | [microsoft/mimalloc](https://github.com/microsoft/mimalloc) |
| TypeScript | 隔离编译器的类型检查与编译，以及 API 声明生成工具 | [microsoft/TypeScript](https://github.com/microsoft/TypeScript) |
| Office Scripts API Reference | 脚本类型声明的固定定义来源，行为由 QuickExcel 独立实现 | [OfficeDev/office-scripts-docs-reference](https://github.com/OfficeDev/office-scripts-docs-reference) |
| Monaco Editor | 脚本编辑器 | [microsoft/monaco-editor](https://github.com/microsoft/monaco-editor) |
| PDF.js | PDF 预览 | [mozilla/pdf.js](https://github.com/mozilla/pdf.js) |
| esbuild | 本地 Web 资源构建工具 | [evanw/esbuild](https://github.com/evanw/esbuild) |
| .NET Deployment Tools | Microsoft.DotNet.Mage 与 ClickOnce launcher | [dotnet/deployment-tools](https://github.com/dotnet/deployment-tools) |
| Open XML SDK | 开发工具中的 XLSX / OOXML 兼容性校验 | [dotnet/Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK) |
| xUnit.net | 领域、存储、计算、脚本与渲染测试框架 | [xunit/xunit](https://github.com/xunit/xunit) |
| xUnit Visual Studio Runner | xUnit 的 Visual Studio / `dotnet test` 适配器 | [xunit/visualstudio.xunit](https://github.com/xunit/visualstudio.xunit) |
| Visual Studio Test Platform | Microsoft.NET.Test.Sdk 与测试执行基础设施 | [microsoft/vstest](https://github.com/microsoft/vstest) |
| QuickJS-NG | 历史脚本引擎，已由 JSC 替代，仅保留来源与回退记录 | [quickjs-ng/quickjs](https://github.com/quickjs-ng/quickjs) |
| WriteableBitmapEx | 历史 WPF 网格绘图依赖，仅保留许可证与来源归档 | [reneschulte/WriteableBitmapEx](https://github.com/reneschulte/WriteableBitmapEx) |
