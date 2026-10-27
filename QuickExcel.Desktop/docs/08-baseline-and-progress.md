# 08 · 原型基线与实施记录

记录日期：2026-10-22。第 1–5 节保留[实施计划](05-implementation-plan.md)首切片「基线与契约」的历史证据，包括最终 102 例；第 6 节追加 0.2 M1 实现与验证。历史首切片状态不代表当前最新状态，也不改写成新版本已经通过的证据。

## 1. 源码、环境与发布模式

本轮起点是 QuickExcel 提交 `e0dbd018db6275be6cec20894c447dcb9d7cef25`，初始工作树无改动。ExcelIO 是同级独立依赖仓库，本机读取到的提交是 `e4534c577cd0ba2162965fb8a9061e081af63c49`；其未跟踪 `.claude/` 不属于 QuickExcel 源码或本轮变更。本轮不修改、复制或提交 ExcelIO 源码。

| 项目 | 2026-10-22 本机读取结果或源码契约 |
| --- | --- |
| 系统 | Windows NT `10.0.22621.0`，x64；开发机检查，不是质量专题中提议的固定性能基准机 |
| SDK / Runtime | .NET SDK `10.0.401`；已安装 `Microsoft.NETCore.App` 与 `Microsoft.WindowsDesktop.App` `10.0.12` |
| 命令环境 | PowerShell `7.6.1`；CMake `4.3.0-rc1`；Visual Studio C++ 安装版本 `18.10.12217.157` |
| 原生引擎 | QuickJS-NG `0.17.0`，来源、许可证及下载校验见 [SOURCES.md](../../QuickExcel.Scripting/native/SOURCES.md) |
| 原生构建契约 | Visual Studio generator，x64，Release，目标 `quickexcel-js`；桥接层 C11，静态 MSVC CRT；关闭安装和示例；链接核心 `qjs`，不引入 `quickjs-libc` |
| 托管构建输出 | 始终指定 `--artifacts-path .\artifacts\build`，包括同级项目引用的输出 |
| 发布模式 | Release x64、`--self-contained false`，需要 .NET 10 Windows Desktop Runtime；复制完整发布目录 |

`dotnet --info` 在沙箱读取 workload 信息时抛出 `InstallerBase` 的 `TypeInitializationException`。SDK 与 runtime 信息分别由 `dotnet --version`、`dotnet --list-runtimes` 读取成功；此异常不等同于构建失败，也不要求安装 WASM workload。完整命令输出仅留在被忽略的 artifacts 中，不把包含个人路径的原始环境输出提交到 Git。

本轮复用已有 native DLL，没有宣称强制重建。普通项目构建只在 DLL 缺失时触发 native 构建，以下哈希只证明本次观察到的文件身份：

| 文件 | SHA-256 |
| --- | --- |
| 下载的 `quickjs-ng-v0.17.0.zip` | `9DE1BF573606D2AEBBBAB6217E616A66E09E3613518489B37B9B3157CF398DCC`，与源码 pin 一致 |
| `QuickExcel.Scripting/native/bin/quickexcel-js.dll` | `C8E37A43CFE755E0FDB9FE5F2C29B7361CEBAB3713BEF8275B70BEAF9DF78DA1` |
| 既有发布目录中的 `quickexcel-js.dll` | 与上行一致；文件修改时间为 `2026-10-22T01:10:14.7192314Z` |

原生源码没有变化时可复用这份已标识的缓存做原型回归。QA-SUP-01 的干净源码重建、编译器缓存键与最终发布包来源证明仍待后续交付流水线实现。

## 2. 本轮验证证据

本轮验证目录为 `artifacts/validation/m1-dom001`。日志与 TRX 不提交到 Git；摘要、命令及限制写入本页。构建和测试由单一执行者顺序运行，避免共用 `artifacts/build` 时的输出竞争。

| 检查 | 当前结果 | 本地证据 |
| --- | --- | --- |
| 变更前 Release solution 构建 | 通过，0 警告、0 错误 | `build.log`，`2026-10-22T02:16:32Z` |
| 原型 Core 基线 | 21 通过、0 失败、0 跳过 | `core-baseline.log`、`core-baseline.trx`，`2026-10-22T02:16:42Z` |
| 新领域契约后的 Release solution 构建 | 通过，0 警告、0 错误 | `build-domain.log` |
| 新领域契约后的 Core 测试 | 101 通过、0 失败、0 跳过；含 80 个领域案例与原有 21 个原型案例 | `core-domain.log`、`core-domain.trx` |
| 日期极值回归后的最终 Core 测试 | 102 通过、0 失败、0 跳过；含 81 个领域案例与原有 21 个原型案例 | `core-domain-final.log`、`core-domain-final.trx`；之后仅更新文档，不改变运行时代码 |
| native smoke | PASS；10,000 次调用单次样本 135 ms | `native-smoke.log`；复用已标识 DLL，不声称本轮原生重建或性能门禁通过 |
| 默认发布脚本 | Core 测试 101 通过；向既有 `artifacts/QuickExcel` 复制若干 DLL/PDB 被拒绝，命令失败 | `publish.log`，保留失败结果；未修改 ACL、终止已有进程或覆盖受保护文件 |
| 独立目录发布 | 相同 Release、framework-dependent 参数发布成功 | `publish-isolated.log`；产物在本轮验证目录的 `publish` 子目录 |
| WPF smoke | 新发布程序退出码 0，新 `smoke-result.txt` 为 PASS，XLSQ/XLSX/示例/PNG 齐全 | 新目录路径记录在 `ui-directory.txt`，目录内有 `exit-code.txt` |

原型基线命令及本轮最终复验入口：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal --logger "trx;LogFileName=core-baseline.trx" --results-directory .\artifacts\validation\m1-dom001
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1
```

最终复验使用 `core-domain-final.trx`，保留此前的 `core-domain.trx` 与变更前的 `core-baseline.trx`。默认发布目录拒绝复制后的替代命令如下；默认脚本的失败仍留在上表，不计作成功。WPF smoke 按[质量专题](06-quality-performance-and-delivery.md)使用唯一新目录运行，保存进程退出码，验证新生成的 `smoke-result.txt`，没有读取旧目录冒充本轮结果。最终增加的日期极值案例只修改测试文件，WPF smoke 运行的领域实现与最终实现相同。

```powershell
dotnet publish .\QuickExcel.Desktop\QuickExcel.Desktop.csproj -c Release --artifacts-path .\artifacts\build --self-contained false -o .\artifacts\validation\m1-dom001\publish -v minimal
```

最后执行只采集、不构建或运行工作簿的脚本：

```powershell
.\scripts\record-baseline.ps1 -EvidenceDirectory .\artifacts\validation\m1-dom001
```

[record-baseline.ps1](../../scripts/record-baseline.ps1) 在该目录下新建唯一 `baseline-capture-*` 子目录，记录环境、两个仓库的 HEAD/status、本仓库源码 SHA-256、相对 HEAD 的二进制 patch、本仓库未跟踪源码副本、native 文件哈希及已有验证产物哈希。输出严格位于 QuickExcel 的 artifacts 内；ExcelIO 只读身份和状态，不复制其源码。每份元数据明确说明采集本身不代表任何构建或测试通过；新鲜运行结果由日志、TRX、命令时间和进程退出码共同证明。采集期间应暂停修改源码，最终文档更新后重新采集一次。

## 3. 固定的 21 个 Core 原型案例

来源：[WorkbookDocumentTests.cs](../../QuickExcel.Core.Test/WorkbookDocumentTests.cs)。原型包含 19 个 Fact 和一个包含两组 InlineData 的 Theory，合计 21 个案例。未来增加领域测试时，保留这组命名案例作为 M0 回归集，不再把整个项目统称为「21 个测试」。测试类为每例创建独立临时目录，完成后删除文件；下面的源码和断言是可复现夹具，临时二进制文件不是持久归档。

| # | 用例名 | 固定行为或夹具 |
| --- | --- | --- |
| 1 | `BlankDocumentSupportsLiteralsAndFormulaEditing` | A2=`10`，A1=`=A2*3` 得到 `30`；前置单引号可输入公式形状的文本 |
| 2 | `CalculationsFollowDependenciesAndBoundFullColumnRanges` | C1→B1→A1 依赖和整列 `SUM(A)` 得到 `16` |
| 3 | `MalformedFormulaDoesNotPreventOtherCellsCalculating` | `=$` 得到 `#VALUE!`，相邻 `=2+3` 仍得到 `5` |
| 4 | `CircularDependenciesAreVisibleInsteadOfUsingStaleNumbers` | A1/B1 循环及依赖循环的 C1 显示 `#CYCLE!`，保留诊断 |
| 5 | `UnsupportedCrossSheetReferenceRetainsItsSavedResult` | `=Sheet2!A1` 保留缓存 `42` 并提示不支持跨表 |
| 6 | `FormulaNumbersUseInvariantCultureAndRestoreTheCallerCulture` | `de-DE` 调用环境下 `1.5+0.25` 得到 `1.75`，调用方区域设置恢复 |
| 7 | `DemoKeepsUsefulCachedValuesUntilScriptsAreEnabled` | 演示合计缓存 `1066.5`；停用期间编辑不伪造新合计，启用后金额 `180`、合计 `708.3` |
| 8 | `JavascriptHandlesScalarAndRangeArgumentsAndRecompiles` | DOUBLE 标量和 TOTAL 一维区域分别得到 `6`、`7`；源码变化停用旧 runtime，重编译后 `9` / `#NAME?` |
| 9 | `JavascriptErrorsProduceCellErrorAndDiagnostic` | JS 抛错得到 `#SCRIPT!` 和可查询诊断 |
| 10 | `CompilationFailureDisablesScriptsAndPreservesCachedResults` | 脚本载入失败后停用，演示合计缓存 `1066.5` 不丢失 |
| 11 | `PackageRoundTripPreservesSourceFormulaCacheStyleAndSheetsWithoutExecutingScript` | XLSQ v1 保存源码、公式、缓存、粗体、第二张中文工作表；打开不执行故意抛错的脚本 |
| 12 | `ExportWritesOnlyValuesWithoutChangingDocumentFormulas` | XLSX 导出只含值 `1066.5`；原文档 JS 与 SUM 公式保持 |
| 13 | `SparseRowsRetainOriginalColumnCoordinatesThroughBothFormats` | E1=`far right`、B4=`=6*7`；XLSQ/XLSX 往返坐标保持、A1/A4 为空、B4 导出 `42` |
| 14 | `ImportedExcelFormulaAlwaysDisplaysLeadingEqualsInEditor` | XLSX 原始公式 `6*7` 进入编辑器显示 `=6*7`，重新提交得到 `42` |
| 15 | `SavingExistingPackageReplacesItAndFailedSaveKeepsPreviousFile` | 成功覆盖到 `second` 后，超限脚本保存失败仍保留上次文件，不遗留 `.tmp` |
| 16 | `InvalidVersionOrMissingScriptFailsClearly`（版本） | `formatVersion:2` 对 v1 读取器明确失败 |
| 17 | `InvalidVersionOrMissingScriptFailsClearly`（脚本） | v1 缺少约定脚本条目明确失败 |
| 18 | `OversizeScriptAndDuplicateEntriesAreRejected` | 重复 manifest 和超过 1 MiB 的脚本被拒绝 |
| 19 | `ImportRejectsRowsThatWouldForceHugeDenseAllocations` | 包含第 1,048,576 行的最小 XLSX 片段在分配前拒绝 |
| 20 | `EmptyImportedWorkbookGetsOneEditableSheet` | 空 XLSX 打开后得到一个可编辑工作表 |
| 21 | `DocumentsDoNotAccumulateGlobalCellChangeHandlers` | 反复创建/停用/释放 20 个文档不累积全局事件处理器 |

## 4. native、WPF 与归档文件

[native smoke](../../QuickExcel.Scripting/smoke/Program.cs) 的固定脚本覆盖大小写函数名、数字和一维区域参数、中文字符串、布尔和 null、console 日志、宿主 API 缺失、对象/Promise 拒绝、递归栈限制、无限循环中断、中断后恢复、跨线程串行访问、非法函数名、载入超时和 80 MiB 分配失败。最后记录 10,000 次 DISCOUNT 调用耗时；它是单次探索样本，不是 P06 的 p95 或 worker 性能结论。

[WPF smoke](../MainWindow.xaml.cs) 固定检查 DOUBLE 计算 `42`→公式栏编辑 `50`→撤销 `42`→重做 `50`；停用脚本后撤销恢复公式缓存；非法粘贴不部分提交；活动 DataGrid 编辑 `30` 在保存前提交，使 B1 缓存为 `60`；XLSQ 重开不自动启用脚本，显式启用后仍为 `60`，随后导出值快照并渲染示例。

本轮开始前已经存在 `artifacts/migration-smoke-74797fb22b2444469d6740947e294bfb`。其中 `smoke-result.txt` 写有 PASS，但未关联本轮执行的进程退出码；这些文件只作历史可检查样例，不能为新源码背书。2026-10-22 读取的哈希如下：

| 历史文件 | SHA-256 |
| --- | --- |
| `QuickExcel-示例.xlsq` | `E7931F226F2537B9B445128C9AACC44A469A6603EB32B3623017575FEAFA52C7` |
| `smoke.xlsq` | `304877CD9C5AC24EA9DC7A40D8582F899508911B2739BE6FB15D1142A4AB6ED5` |
| `smoke.xlsx` | `AD9F187B7DAB1EB82E35E15918FE67EF08E86EE5F267013C01BC5681780B58BA` |

新 smoke 输出要与本轮日志一起留在验证目录，再由采集脚本记录哈希。ZIP 容器时间戳可能导致相同逻辑夹具每次生成的文件哈希不同；哈希用于标识某次产物，跨次等价性由工作表地址、公式、缓存和类型断言确定。

## 5. 历史首切片的能力边界与状态

| 范围 | M0 已有行为 | 后续目标及本轮边界 |
| --- | --- | --- |
| 值与类型 | 原型以字符串存储，公式与显示结果有现有字段 | DOM-001 引入独立类型契约；尚不等于 XLSX 适配、日期格式或 UI 类型迁移完成 |
| 稀疏性 | 现有回归确保稀疏坐标正确 | 底层仍可能扩展稠密列表；DOM-002 才保证存储随实际单元格数量增长 |
| 计算 | 同表依赖、循环/错误诊断、脚本启停与缓存 | 跨表、typed evaluator、增量重算、完整新鲜度状态和过期结果门禁待 M1/M3 |
| 文件 | XLSQ v1、显式启用脚本、值快照导出、部分输入预算、临时文件替换 | XLSQ v2、完整类型保真、未知部件保留/损失报告与故障恢复待 M2/M6 |
| 脚本 | 每文档进程内 QuickJS；64 MiB 堆、256 KiB 栈、约 1 秒中断期限 | 无独立 worker 或 OS 沙箱；约 5 秒重算检查不是硬截止；M4/M6 才验收隔离 |
| UI | 公式栏、单元格编辑、粘贴及内容撤销、保存编辑中的值 | 尚未迁移统一领域会话/事务；真实键盘、IME、无障碍不由 smoke 全部覆盖 |
| 容量 | Core 100,000 行、1,024 列、1,000,000 个计入限制的单元格；UI 256 列、粘贴 20,000 格 | 当前部分空洞计入容量；模型容量不等于 UI 可编辑范围或性能预算已达成 |

| 任务 | 本轮状态 | 完成判据与剩余工作 |
| --- | --- | --- |
| BASE-001 | 完成 | 源码起点、环境、来源、缓存 DLL 与发布模式已记录；采集脚本冻结本轮工作树与产物清单，native 未重建和默认发布目录拒绝保留为明确限制 |
| BASE-002 | 完成 | 21 例源码夹具、演示/启停/稀疏/包/导出断言已固定；本轮新 WPF 文件归档目录由 `ui-directory.txt` 定位 |
| BASE-003 | 完成 | 当前/目标能力、21 原型案例、native/WPF 证据和默认发布失败已逐项分开记录 |
| DOM-001 | 完成 | 不可变值、内容、独立缓存状态和日期辅助契约通过 81 个领域案例，见 [DomainValueTests.cs](../../QuickExcel.Core.Test/DomainValueTests.cs) 与 [ADR-014](07-architecture-decisions.md#adr-014dom-001-的值缓存与日期边界)；未改变现有 facade 路径 |
| DOM-002、APP/CMD、M2–M6 | 待开始 | 本轮不以值契约或现有 WPF 行为代替这些任务的退出门禁 |

未执行的 ExcelIO 独立测试、原生干净重建、故障注入、固定性能基准和无 SDK 干净机器验收，均不得记作通过。本轮不把 M1 整体标为完成。

## 6. 0.2 M1 试用版实现与验收

本次按用户确认交付 0.2 试用版，范围是 M1，不扩展为 M2–M6 完整 Beta。Git HEAD 仍为 `e0dbd018db6275be6cec20894c447dcb9d7cef25`，新领域、会话、适配、桌面与文档代码存在于**未提交工作树**；HEAD 单独不能代表当前源码。没有修改或复制同级 ExcelIO 源码，未创建 PR。

### 6.1 当前编译边界与实现

| 任务 | 0.2 实现 | 回归与边界 |
| --- | --- | --- |
| DOM-002 | ImmutableDictionary 稀疏格、会话内稳定 SheetId、不可变样式/尺寸/图片快照 | [WorkbookSnapshotTests.cs](../../QuickExcel.Core.Test/WorkbookSnapshotTests.cs) 覆盖远端单格、随机参考模型、旧快照、改名/重排身份；领域改名/重排验证不等于桌面已有对应命令 |
| APP-001 | DocumentSession 串行写入门禁、Revision、SavedRevision、关闭生命周期和统一通知 | [DocumentSessionTests.cs](../../QuickExcel.Core.Test/DocumentSessionTests.cs) 验证事务/通知/关闭；无同时写入的旧可变 Workbook 路径 |
| CMD-001/002 | Set/Paste/Clear 全量预检，内容和新增表撤销/重做，历史预算默认 32 MiB | 非法末项/容量/重复地址不产生半次提交；撤销恢复类型、公式和旧缓存，修订继续递增；历史截断在界面可见 |
| APP-002 | 输入快照与四元组提交门禁：sessionId、inputRevision、calculationGeneration、scriptGeneration | [WorkbookFacadeTests.cs](../../QuickExcel.Core.Test/WorkbookFacadeTests.cs) 用可控端口覆盖乱序、关闭/重开、源码/授权变化、取消后返回；没有伪造 worker/request 身份 |
| APP-003 | Core 仅依赖 .NET；新增 Infrastructure 承载包封装、ExcelIO 计算/存储和 QuickJS factory；DesktopComposition 注入 | [Core 项目](../../QuickExcel.Core/QuickExcel.Core.csproj) 无 ProjectReference；[Infrastructure](../../QuickExcel.Infrastructure/QuickExcel.Infrastructure.csproj) 才依赖 ExcelIO 与 Scripting |
| 旧存储适配 | 继续 XLSQ v1；元数据和图片深复制；读取保留显式稀疏坐标；旧 IO 分配有总量预算 | [StorageAdapterTests.cs](../../QuickExcel.Core.Test/StorageAdapterTests.cs) 覆盖历史 Text、缓存 Stale、样式/图片、导出、超限和被占用文件替换失败 |
| UI-001 | EditorController、最小 ViewModel、统一编辑/撤销/保存入口；删除窗口自有 dirty 与撤销栈 | WPF 新鲜 smoke 已通过，具体目录见下表；源码编辑撤销授权，重算失败仍允许保存有效输入和旧缓存 |

保存从 captured Snapshot 开始，成功后只标记该修订已保存，IO 期间的新编辑仍 dirty；多个保存串行捕获，失败不推进 SavedRevision。该基础在 M1 提前交付，M2 的 v2 格式和完整保存协议仍未完成。Core/Infrastructure 分离是实际程序集边界，不只是命名空间整理。

当前计算适配基于类型值读取快照，每轮重建 QuickJS runtime；同表兼容计算继续通过既有行为回归。跨表公式及传递依赖保留 Blocked 缓存，不宣称跨表或完整增量计算。运行时仍在进程内，原生调用不能靠托管 CancellationToken 在调用中强制终止；可信脚本边界保持不变。

### 6.2 最终验证与发布证据

当前验证目录为 `artifacts/validation/release-0.2`。以下区分源码构建、桌面 smoke 和最终候选包验证，不把不同阶段混成一次发布成功。

| 检查 | 已确认结果 | 本地证据 |
| --- | --- | --- |
| Release solution | 0 警告、0 错误 | `build.log`、`build-final.log`；包含新增 Infrastructure 与纯 Core |
| Core 回归 | 221 通过、0 失败、0 跳过 | `core.log`、`core.trx`；与历史首切片 102 例分别保留 |
| 首次扩展 WPF smoke | 发现同格公式栏→网格编辑切换失败，未计通过 | `ui-3e7d3c370a7c46c2bea21d7a18774f15/smoke-error.txt`，保留失败证据 |
| 修复后扩展 WPF smoke | PASS；主验证流程确认成功退出，新结果文件含扩展场景 | `ui-72b546950ac14798969d4b12b0efd21a/smoke-result.txt`；后续双向切换回归目录由 `ui-directory.txt` 定位 |
| 发布版焦点断言修正 | 前两次未放行；WPF 将焦点移回公式栏时可清空 CurrentCell，原断言误拒绝此状态 | `release.log`、`release-retry.log` 和对应新 smoke-error；修正后要求不存在错误的有效地址，待编辑原地址及修订仍正确 |
| 最终发布包 native/WPF smoke | PASS；Core 221/221，native 10,000 次调用 163 ms，WPF 退出 0 且新 PASS | `release-final.log`；`artifacts/smoke/20261022-030232-c3cbff2e`，包内 metadata 三项 passed |
| ZIP 与 SHA-256 | 生成并核对；21 文件、1,121,398 字节，无用户工作簿/截图/日志 | `archive-validation.json`；包名与哈希见下文 |
| 解压后独立启动 | PASS；新解压目录运行完整 WPF smoke，退出 0 | `extracted-smoke-35afe56bec874306ac5e6604dee2e404/smoke-result.txt`，目录与校验记录在 `archive-validation.json` |
| 干净机器、签名与安装器 | 未执行 / 未实现 | 不属于本次 0.2 已完成范围，仍属 M6 |

通过的扩展 WPF smoke 覆盖：JS 和公式栏计算、公式缓存撤销、非法粘贴原子拒绝、活动网格编辑保存、公式栏未按 Enter 直接保存、同格编辑器双向切换、无效待提交输入保持原目标和既有文件、新增表撤销/重做、保存修订及脚本授权不误标 dirty、无效脚本编辑撤销授权，以及计算失败后撤销/保存/关闭有效输入。它没有替代真实 IME、读屏、全部焦点组合和性能专题。

最终交付 `artifacts/releases/QuickExcel-0.2.0-win-x64-20261022-030232-c3cbff2e.zip`，SHA-256 为 `0EF8D34FEB744F004C0BFC05753C52BCBA552A0358BFBC332C686EA24227C567`。默认启动目录 `artifacts/QuickExcel` 已更新，与 ZIP 内应用 DLL 哈希一致；根目录 `Start-QuickExcel.cmd` 可直接启动本版。源码、未提交 patch 和测试产物由 `record-baseline.ps1` 在本验证目录内另存唯一快照，不修改发布 ZIP。

### 6.3 发布入口与真实性说明

```powershell
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

脚本始终为托管测试、运行和发布指定 `--artifacts-path`；发布使用 `-r win-x64 --self-contained false`。程序需要 .NET 10 Windows Desktop Runtime x64；仅在装有相应 Runtime 的开发环境运行，不声称已通过干净机器安装验收。

每次 publish 与 WPF smoke 使用新的唯一 artifacts 子目录。WPF 使用 Hidden 进程、45 秒看门狗，只能终止本次脚本自行创建的测试进程；同时要求退出码 0 和本次新生成的 PASS 文件。旧输出目录可能含用户工作簿，归档只使用本次全新 publish tree，不将 smoke 数据、截图、日志装入发布包。

ZIP 名称包含版本、UTC 时间与随机后缀，旁边生成 `.sha256`；旧 ZIP 不覆盖。包内 `BUILD-METADATA.json` 记录 HEAD/dirty、构建参数、SDK/Runtime、QuickJS 源码 pin 与实际 DLL SHA-256、各项本次验证状态。工作树未提交和 `-SkipTests` 都如实记录；哈希是产物身份，不是数字签名。

QuickJS 许可证随包提供，ExcelIO 两个 .NET 项目声明 MIT；依赖 notice 只照录可核查元数据，不从其独立 npm 项目推定版权文本。普通构建可能复用已有 native DLL，本次仍不凭一次成功 publish 宣称完成 native 干净重建。

### 6.4 公开限制与后续状态

- M1 实现与核心/WPF 回归完成；最终发布包 smoke、ZIP/hash 和解压复验通过，0.2 试用版已放行。
- v1 的 literal 仍写为文本，重开保留 Text；不猜测恢复数字、布尔、错误或丢失的日期类型，空白/空文本可能合并。内存类型不等于文件保真，SheetId/Revision 不写入 v1。
- 公式缓存可能过期。有效输入的保存不因重算失败被阻断，旧缓存不会被伪装为新结果；无效编辑仍须修正后才保存。运行记录和状态栏明确提示。
- 任意脚本编辑立即撤销授权，包括校验失败的文本；打开/重开/保存不携带运行授权。进程内 JS 仅限可信脚本，没有 worker 或 OS 沙箱。
- M2 类型/日期/v2、M3 跨表/增量、M4 worker/动作、M5 完整交互、M6 恢复/性能/签名与干净机器均保持后续范围。0.2 不使用“完整 Beta”描述。

## 7. 0.2.1 关闭窗口修复

用户在运行 0.2.0 后报告 WPF 异常：“在窗口关闭期间，无法将可见性设置为可见，也无法调用 Show、ShowDialog、Close 或 WindowInteropHelper.EnsureHandle。”根因是 `Window_Closing` 内等待的 `ConfirmDiscardAsync` 在干净文档或“不保存”分支同步完成，于是第二次 `Close()` 在第一次 Closing 尚未退出时重入。第 6 节 smoke 虽验证了确认方法返回值，但最后使用 `_closingApproved = true` 关闭，没有覆盖真实关闭事件；其通过不能证明该路径正确。

修复将整个确认/保存流程放入 Dispatcher 队列，Closing 立即取消并返回。待关闭标记防止重入，取消或失败后清除；文档释放从 Closing 移到实际 Closed。排队自动重算在待关闭及关闭后不执行，取消关闭后可恢复。App 在 smoke 模式下将未处理 Dispatcher 异常写入证据并以非零状态退出，避免错误对话框阻塞无人值守检查。

[MainWindow.CloseSmoke.cs](../MainWindow.CloseSmoke.cs) 使用真实 WPF 子窗口调用 `Close()`，覆盖 8 类场景：干净关闭（重复请求）、丢弃修改且阻止排队重算、取消后继续编辑并重试、非法待编辑内容、保存后关闭、保存失败保留文档与原文件、正在执行工作时拒绝关闭、最终 Closing 被其他处理器取消时不释放会话。测试成功路径不预设关闭批准标记。确认选项采用注入，未自动操作原生另存为对话框。

验证已完成：源码构建 0 警告、0 错误；Core 221/221 通过；native smoke 通过；发布包和 ZIP 解压后的 WPF smoke 均以 0 退出，上述 8 类真实关闭场景及既有编辑/保存回归全部通过。证据位于 `artifacts/validation/release-0.2.1/release.log`、`archive-validation.json`，发布 smoke 目录为 `artifacts/smoke/20261022-031511-ba477dd9`，解压复验目录为 `artifacts/validation/release-0.2.1/extracted-smoke-364b785b62a24c84ab784cdbf3608cd7`；均包含新生成的 `close-smoke-result.txt`。

修复包：`artifacts/releases/QuickExcel-0.2.1-win-x64-20261022-031511-ba477dd9.zip`（1,127,145 字节），SHA-256：`675B1C33E2F88A6BB528304C859805FDAFE2D23A1465328B08837264BA3760AD`。默认 `artifacts/QuickExcel` 已更新，应用 DLL 与 ZIP 内一致；`Start-QuickExcel.cmd` 继续启动这个已更新目录。该证据仅用于本修复版，不改写 0.2.0 历史。

## 8. 0.3 文件预览 / M2（2026-10-27）

从 QuickExcel `82381fe` 继续实现类型化 XLSX/XLSQ v2、无猜测 v1 迁移、稳定身份/修订、模块元数据、文件预算与关系校验、兼容报告和写后重开验证，并接入桌面。源码尚未提交，ExcelIO 源码未改。

Release build 0 警告/0 错误；Core 266/266；native、发布版及最终 ZIP 解压后的 WPF smoke 全部通过。新增 WPF 场景验证真实保存入口的数字/文本/布尔/空文本、v2 manifest、稳定身份和 v1 原件保护；没有自动操作原生文件选择对话框。

最终包及完整支持范围见 [M2 文件与类型](09-m2-files-and-types.md)。证据为 `artifacts/validation/release-0.3-preview`，默认启动目录已更新。后续阶段为 M3 跨表与增量计算；本版不是完整 Beta。

## 9. 0.4 计算预览 / M3（2026-10-27）

在 M2 的 813e91d 上完成支持矩阵内的跨表/名称 typed evaluator、稀疏工作簿依赖和 SCC、增量求值闭包、可撤销结构命令、JS 描述符/二维参数桥及可移植公式导出。桌面组合根已切换；旧适配器和既有回归保留。

最终 Release solution 0 警告/0 错误；Core 331/331（新增 65）；native、发布版 WPF、最终 ZIP 独立解压 WPF 全部通过。默认启动目录已更新。支持范围、不可移植闭包规则、脚本版本选择、实际证据和最终 ZIP/hash 见 [M3 记录](10-m3-calculation.md)。

本轮未创建提交，未修改同级 ExcelIO 源码。下一阶段为 M4 worker/脚本动作；0.4 不等于完整 Beta。