# JSC / TypeScript 实施记录

2026-11-03。对应 [70 号方案](70-jsc-typescript-migration-plan.md)。J0–J6 已完成本轮声明的工程范围：同步模块桥、墙钟中断、隔离 TS 编译、文件/签名、桌面、独立 Windows/Linux 隔离/恢复/发布/性能验证，以及单引擎清理和默认激活。Windows 默认版本已切换到 `0.6.0-preview.8`。**`JSC-MEM-001`：64 MiB 引擎总分配硬限额仍未实现；按用户明确要求延期处理，不作为后续阶段门禁。** 不把 256 MiB Job 或 GC 阈值称为等效实现。源码已改为 JSC 单引擎、IPC 6；下文 J0–J5 中“候选”“默认 QuickJS”等为阶段历史，最终激活与回退见 §11。

## 1 已落地的开发入口

- `tools/QuickExcel.JscProbe/build.ps1`：校验固定来源，源码构建 ICU 和 Windows x64 JSCOnly，生成产品原生库和构建身份；历史探针须显式传入 `-IncludeProbe`。
- `prepare.py`：校验归档及已有源码缓存；不运行安装器，不修改仓库外的工具或库。
- `probe.cpp`：使用 JSC C API 的独立可信测试程序，提供执行限时、异常路径检查和进程内存采样。它可读取本地测试文件，**不是生产 worker 或安全沙箱**。
- `QuickExcel.Scripting/TypeScript/frontend.js`：官方 TS Compiler API 的内存 host，开发探针区分仅转译和完整检查；产品仅完整 Program 检查/生成，保留 ESM，使用固定标准库和产品 `quickexcel.d.ts`。
- `run.py`：功能用例与性能采样分开运行，使用新进程并保存逐次原始结果。超时由外部测试驱动兜底；这不是产品的 Job/AppContainer 证明。
- Windows/Unix 发布脚本显式拒绝 `JscProbe`、`quickexcel-jsc-probe`、`TypeScriptProbe`。J1/J2 已通过显式开发属性接入托管桥和完整 worker payload；默认启动器和发布入口未切换。

复现命令（仓库根目录）：

```powershell
.\tools\QuickExcel.JscProbe\build.ps1 -IncludeProbe
python .\tools\QuickExcel.JscProbe\run.py
python .\tools\QuickExcel.JscProbe\run.py --performance --repeats 3
```

构建/测试产物全部位于忽略的 `artifacts/jsc/`。完整用法与限制见 [探针 README](../../tools/QuickExcel.JscProbe/README.md)。这些命令无需 Node.js/Bun，但开发机需要文档中列出的 C++ 构建工具。

## 2 固定来源与构建差异

| 项目 | 固定候选 |
| --- | --- |
| WebKit | GTK 2.54.2，tag `webkitgtk-2.54.2`，commit `b95c91b1c0ee43cc86d25ff7af73038f67a88e04` |
| 源码归档 | `webkitgtk-2.54.2.tar.xz`，SHA-256 `a76345b391bef6211fb5ab6cfab2e86b285d18cda17cbb9ade595f1252b86065`，与上游 sums 核对 |
| Windows 补充源码 | GTK 归档省略的 WTF `win`/`text/win` 目录共 27 个文件，来自同一 commit；冻结时核对 Git blob，构建时逐文件核对 SHA-256 |
| ICU | 77.1，从源码构建完整 data/i18n/uc DLL；拒绝把 3 KiB 的 data stub 当作完成产物 |
| 编译器 | LLVM clang-cl 20.1.8，MSVC v143 14.44.35207 工具/头文件；具体 SDK、二进制摘要进入构建身份 |
| 生成工具 | RubyInstaller 4.0.9-1、Git for Windows Perl、Python 3.13、CMake 4.3 |
| TypeScript | 5.9.3；归档 SHA-512 已与 npm integrity 核对；与 Monaco 0.57.0 的 TS 依赖一致 |
| JSC 配置 | JIT/DFG 开启，FTL、WebAssembly、remote inspector、API tests 关闭；启用归档内的 mimalloc 3.2.8；尚非产品策略身份 |

所有固定 URL/摘要见 `sources.json` 和 `windows-sources.json`。没有使用第三方预编译 JSC DLL。

构建迭代保留失败日志，并记录四个显式修正：

1. `InlineCacheCompiler.h` 使用 `CCallHelpers` 的嵌套类型，需显式包含其定义。
2. `ScriptExecutable.cpp` 使用 `CompilationResult` 枚举值，需显式包含枚举定义。

3. J1 在 `NodesAnalyzeModule.cpp` 中拒绝全部 import attributes，包括 `with {}`，避免分析器将空属性存在信息抹掉后绕过模块策略。
4. Linux JSCOnly 的 CMake 只在未显式配置时选择 `ENABLE_API_TESTS` 默认值，避免覆盖 `OFF` 后依赖归档未包含的测试源码。

修正位于 `source-adjustments.json`；准备流程验证归档原文后应用，后续拒绝其他源码改动。前两项只增加包含，第三项是明确的模块语法策略调整，第四项仅涉及测试目标开关。另修复构建环境的 Perl PATH 和 Python UTF-8 设置；没有关闭编译错误检查。Linux 使用归档 ICU 的独立 LF 规范化构建副本，保留原始校验缓存；`USE_HEADER_MAPS=OFF` 避免依赖 GTK 归档省略的 hmaptool。

JSC 的 LGPL/BSD、WTF/bmalloc 第三方通知、mimalloc MIT、ICU Unicode 许可、TS Apache-2.0/第三方通知保留在来源树。J5 的 `prepare_distribution.py` 生成完整通知、对应来源归档及自有桥/构建脚本和动态替换重建说明；发布为 `JSC-Licenses/`，不混入私有 worker payload。

最初的 libpas 候选在简单脚本上提交内存约 282 MiB，核验过的 256 MiB Job 中甚至不能完成上下文启动。缩小 JIT 预留没有消除此问题，不能把提交内存都归因于 JIT。`USE_SYSTEM_MALLOC=ON` 被上游 Windows 明确拒绝；改用上游自带的 `USE_MIMALLOC=ON` 后，同类脚本提交内存约 20.6 MiB，可以在 256 MiB Job 中启动。未修改分配器实现来规避限制。旧候选与失败日志保留在 `artifacts/jsc/candidates/libpas/`、`artifacts/jsc/logs/`。

## 3 编译实验的明确边界

J0 候选选项为 ES2022 / ESNext ESM、strict、isolatedModules、verbatimModuleSyntax、类字段 define 语义、完整类型检查成功才生成。API v1 仍是原 JavaScript 语义；J0 当时不改变文件格式，后续 J3 格式见 §8。

内存 host 只解析已声明的精确逻辑 ID，不接受相对路径、外部包、引用指令、动态导入、import attributes、顶层 await 或装饰器；没有 DOM/Node 类型。枚举、namespace、泛型、类字段、仅类型导入和 JS/TS 混合通过专门用例检查。编译输入含顶层 `throw` 的成功用例用于确认编译不执行用户代码。

原文总量限制为 1 MiB UTF-8、最多 64 模块/入口，生成 JS 仍限制为 1 MiB，Source Map 上限 4 MiB；失败不返回部分产物。诊断最多 128 条、紧凑 JSON 共 16 KiB，超限有明确省略标记。诊断行列使用一基，映射的 sources 归一为逻辑模块 ID。J0 不覆盖生产编译 IPC；J3/J4 实现见 §8–9。

根据下述测量，编译默认策略选定为**应用时完整 Program 检查和生成**，仅转译保留为开发对照项，不作为类型错误后的回退。

编译候选预算固定为 768 MiB 进程 Job、5 秒单调墙钟截止、宿主 6 秒强杀兜底、5 MiB JSC 栈；原文/输出/映射/诊断额度如上。512 MiB 在密集 1 MiB 输入的两种前端模式下均失败，不能选作已经验证可用的预算。最早使用的 30 秒只是探索上限，已不作为默认。J0 开发采样严格单进程顺序运行；产品编译 AppContainer、取消和生命周期另有 §8 证据。

以上预算**只用于独立编译角色**。执行仍保持 256 MiB 进程、256 KiB 栈、1 秒调用/5 秒整轮；64 MiB 引擎硬限额以开放 issue 延期，不宣称已实现。

## 4 实测结果和资源差异

机器为 Windows 11 x64（10.0.22621）、Core i9-13980HX、32 逻辑处理器、约 32 GiB RAM，Python 3.13.6。每个样本重新启动探针进程，固定源码与构建身份，顺序采样。下表是进程启动至退出的墙钟中位数和范围，不是仅编译器内部耗时，更不是工作簿整轮计算收益。每种输入/模式仅 3 次，不报告 p95。

| 输入 | 仅转译 ms：中位数 [范围] | 完整检查 ms：中位数 [范围] | 峰值提交 MiB：转译 / 完整检查 |
| --- | --- | --- | --- |
| 小函数 | 305 [299–432] | 487 [485–505] | 121.2 / 239.2 |
| 128 KiB，主要为注释 | 281 [270–379] | 524 [489–550] | 134.2 / 241.0 |
| 接近 1 MiB，主要为注释 | 382 [340–390] | 525 [511–564] | 159.1 / 261.9 |
| 64 个小模块 | 362 [342–386] | 538 [506–540] | 166.7 / 223.4 |
| 128 KiB 密集声明 | 410 [390–438] | 579 [569–620] | 226.2 / 309.9 |
| 接近 1 MiB 密集声明 | 977 [948–1009] | 1172 [1076–1178] | 676.4 / 591.6 |

阶段交互目标据此选为：本机固定小函数冷检查 ≤ 1 秒、密集 1 MiB 冷检查 ≤ 2 秒；后续桌面 IPC、沙箱启动、排队和取消开销需另外验证，不能宣称已经达到端到端目标。注释填充输入不能代表相同字节数的实际 AST 工作量。提交内存不是工作集，也不是引擎总分配计数。

31 项功能/观察用例（10 个原生行为、21 个 TS 用例）得到预期结果，包含结构化诊断、失败无部分产物、Source Map 与类型导入检查。这里的“预期结果”也包含**不满足生产契约的反例**，因此所有 summary 明确写入 `migrationReady: false`。

| 约束 | 实际证据 | 判定 |
| --- | --- | --- |
| 256 MiB 执行进程 | 自分配 Job 后查询限额和成员关系；简单脚本成功，300 MiB 数组请求失败 | 原生开发进程证据通过；不是 AppContainer/生产 launcher 验收 |
| 64 MiB 引擎总分配 | 读回 `gcMaxHeapSize=67108864`，在 256 MiB Job 内仍成功分配并填满 96 MiB TypedArray | **JSC-MEM-001 开放，用户延期，不阻止推进** |
| 256 KiB JS 栈 | 读回 `maxPerThreadStackUsage=262144`，递归产生异常而不是成功返回 | 单一递归路径通过，非所有桥接入口证明 |
| 1 秒中断 | 原始 JSC CPU 看门狗有约 2.1–5.2 秒墙钟样本；单调截止加 10 ms 原生重设后循环/微任务约 1.03–1.74 秒 | 改善但未证明 1 秒契约；需要生产宿主独立截止/终止与完整入口验证 |
| 异常格式化 | 不读取异常对象的 stack/message/toString；恶意 getter 没有运行机会 | 仅探针异常路径，完整值桥尚未实现 |
| 同步初始化 | 求值后用户 Promise 微任务已运行；上游 `JSLock::willReleaseLock` 调用 `drainMicrotasks` | **公开 C API 路径不满足现有同步模块契约** |
| 内存 ESM | TS 输出保留 ESM，但原生探针没有内存 ESM 加载器 | 未实施；不以脚本求值或拼接代替模块 |

固定源码的 `Heap.cpp` 中 `gcMaxHeapSize` 决定是否请求 GC，并非分配前拒绝条件；`maxSingleAllocationSize` 还是仅调试配置的单次分配检查，不能成为 Release 的总额度。当前公开嵌入路径尚无覆盖 GC 对象、外部 buffer、JIT 和引擎原生分配的 64 MiB 额度实现或等效证明。进程提交采样、GC 阈值及 256 MiB Job 均不能替代它。

## 5 原始证据和复现状态

所有目录相对于仓库根目录，均被 Git 忽略，保留失败证据：

| 证据目录 | 内容 |
| --- | --- |
| `artifacts/jsc/logs/20261103-060539-8820de2b/` | 最后一次固定来源复核与构建日志；`build/win-x64/identity.json` 记录来源、选项、SDK、探针源码及实际运行的 EXE/DLL 摘要 |
| `artifacts/jsc/results/20261103-061135-bedb5a5a/` | 最终 31 项功能/行为检查，失败列表为空；执行 256 MiB/256 KiB，编译 768 MiB/5 MiB；包含上述 64 MiB 反例 |
| `artifacts/jsc/results/20261103-060800-c61da2ea/` | 768 MiB 编译预算下的 36 次性能原始样本，失败列表为空；上表来源 |
| `artifacts/jsc/results/20261103-060704-a2c155c0/` | 512 MiB 编译预算失败：6 个密集 1 MiB 样本进程失败，其他 30 个完成 |
| `artifacts/jsc/results/20261103-055206-376476aa/` | libpas 候选的 256 MiB 启动失败和初始行为观察 |
| `artifacts/jsc/results/20261103-060620-e7ac225c/` | 包含约 1.74 秒微任务中断响应的样本；不能只选择后来的较快样本宣称通过 |
| `artifacts/jsc/results/20261103-060930-2e2525a4/` | 包含原始 CPU 看门狗约 5.15 秒墙钟响应的样本 |
| `artifacts/jsc/logs/build-system-malloc-rejected.log` | 上游不支持 Windows system malloc 的失败日志 |

每个结果目录保存实际输入、逐项退出码、JSC 选项读回、墙钟/内部耗时、峰值工作集和提交内存、构建身份及前端/标准库摘要。测试驱动在运行前验证 EXE/DLL 和原生探针源码身份，并拒绝采样期间发生的构建身份变化。下载归档和已有源码缓存已再次校验；新构建日志使用唯一目录，避免覆盖失败记录。

上述是 J0 历史证据，不覆盖后续桥与托管实现。没有运行全仓 .NET 回归或激活新发布包；J1/J2 定向证据见下节。发布脚本静态检查不代替 J5 发布/部署矩阵。

## 6 门禁结论与剩余工作

用户已调整 70 号文档 §4.3：64 MiB 引擎硬限额保留开放 issue，不作为后续推进门禁；其余预算与平台验收仍须满足。本节记录早期阶段结论；当前 J5 已完成，J6 最终状态见 §11。

`JSC-MEM-001` 后续需覆盖 GC 对象、外部 buffer、JIT 和引擎原生分配及失败清理。当前优先按 J3–J6 推进，不因该 issue 停工。

J3/J4 的 Windows 实现与定向证据见 §8–9；J5 两平台独立结果见 §10，J6 见 §11。不能从一个平台推断另一平台通过；各平台保留原方案退出条件。

## 7 J1/J2 同步桥与 Windows Worker

自有 `QuickExcel.Scripting/native/jsc/bridge.cpp` 编译进固定 JSC DLL，仅向托管层导出 C ABI。`JSModuleLoader::makeModule` 只解析并得到已完成的内部 Promise；桥直接构造 `loadedModules`、链接原生模块记录并同步求值。没有执行用户微任务、拼接源码或改写 import。循环、活绑定、TDZ、共享模块一次初始化、入口顺序和不可达模块均有验证。

VM 全生命周期保持 `DrainMicrotaskDelayScope`；v2 在每个入口与调用完成时拒绝待执行任务和动态导入，销毁前清空队列。API v1 保留原有“任务不推进”的行为。错误只读取引擎内部异常数据和原始逻辑源码位置，不调用用户 `stack/message/toString` getter。脚本伪造资源错误文本不会制造原生失效标记。

每个执行操作共用一个 1 秒单调墙钟截止，覆盖模块加载、调用、序列化与异常路径。独立原生线程调用 JSC 并发安全的不可捕获终止 trap，预先创建 termination exception；完成与取消在互斥锁内按代次隔离，避免迟到 trap 污染下一调用。中断、栈溢出、OOM 或 v2 待执行任务使 VM 永久失效；Worker 拒绝整个批次，宿主杀 Job，不复用或重放。宿主原有 5 秒整轮截止及用户取消独立生效。

候选身份为 `JavaScriptCore/2.54.2/b95c91b1/qx-sync-1`，IPC 6，执行策略 `conservative-jsc-v1-memory-issue-open`。默认 IPC 5/QuickJS 构建保持不变；候选不含自动回退。公式、动作、事件及虚拟文件复用同一原生桥，AppContainer、创建时 Job、零能力令牌、只读 payload、句柄白名单及 journaled lease 保留。

- `artifacts/jsc/bridge-results/20261103-064034-18273528/`：24 项桥检查通过；循环调用约 1000 ms 中断，包含捕获中断、后续入口循环、序列化循环和拒绝复用。
- `artifacts/jsc/test-results/jsc-modules.trx`：58 项 API v1/v2 与模块检查通过。
- `artifacts/jsc/test-results/jsc-worker-initial.trx`：首次扩展检查 204 通过、4 失败。保留 64 MiB 已知缺口、旧异常格式/尾调用断言差异及测试 TEXT 名称冲突证据。
- `artifacts/jsc/test-results/jsc-worker-corrected.trx`：修正后相关 36 项通过，1 项明确以 `JSC-MEM-001` 延期。增加非尾递归栈限制、不可捕获终止、异常 getter 不执行、伪造错误文本和待执行任务导致整批作废的回归；并非删除失败来声称等效。

扩展检查覆盖真实 Windows worker、批次、actions、事件/虚拟文件、协议、取消、故障注入、AppContainer 负向访问及崩溃恢复。没有运行全仓回归。这些本机证据不代替 J5 干净机器、完整部署矩阵、Linux 或物理输入验证。

macOS 脚本执行继续关闭；既有 M4/M6、物理输入、打印和干净设备门禁不因本实验改变。

## 8 J3 隔离 TypeScript、格式与身份

固定 TS 5.9.3、标准库、声明和前端被打包为哈希校验的 `typescript.bundle.json`（9,969,983 字节，SHA-256 `36A207D75C76C6B2CE9383BEC163D0FDB2627EDAAA131B8CA2E625A414CBC5A1`）。编译角色通过同一 AppContainer/只读 payload 启动机制运行，同一固定 JSC 原生库；不执行工作簿源码，不暴露文件 host，不启动子执行进程。

角色在握手前由进程启动参数固定。`compile` 握手绑定前端摘要、API 2 和独立策略，J3 首版为 `typescript-check-768m-5s-v1`，J5 经真实进程测量更新为 `typescript-check-1024m-5s-v1`（见 §10）。原生 VM 首次创建后也锁定角色，不能切换执行/编译栈预算。宿主全局并发为 1、队列为 0，每个 worker 只处理一次快照。取消、关闭或 6 秒宿主截止终止独立 Job。编译结束后再开始原有执行轮截止，不挤占或扩大执行预算。

源码先冻结为不可变快照。摘要覆盖排序后的 ID/语言/原文、入口顺序、选项、前端、声明、引擎及策略。成功派生缓存按源码集合持有，至多 4 个条目/16 MiB 序列化数据；失败和取消不入缓存。返回结果校验身份、输出摘要、条数、总字节和 Source Map 结构/VLQ，拒绝外部路径及伪造映射。运行错误转换为原始逻辑 ID 和一基 UTF-16 行列；不能映射的位置明确标为生成 JS。

含 TS 的 XLSQ 要求 `typescript-sources` 特性、明确模块语言、固定选项 `ts5.9.3-es2022-esnext-strict-isolated-verbatim-define-v1` 和 API 2。只保存原文，生成 JS/映射不写入包。TS 使用签名载荷 v2，绑定语言和选项；纯 JS 保持原 v1 签名载荷字节格式。语言变化同样撤销授权/签名并推进脚本修订。脚本源码仍沿用既有独立编辑历史，不额外塞入工作簿单元格撤销栈。

新 JSC 策略下，旧证书签名继续验证，但不直接根据旧证书 pin 自动执行打开事件；桌面要求显式应用/授权。未删除签名或悄悄把旧策略授权当作新授权。

证据：`artifacts/jsc/test-results/jsc-ts-storage.trx` 首轮 89 通过、1 个脚本编辑历史边界断言失败；修正错误测试预期后 `jsc-ts-contracts.trx` 31 通过。覆盖跨模块检查、顶层代码不运行、类型失败、危险导入、取消、错配输出/映射拒绝、原始错误位置及原文/语言/特性/签名保存重开。失败记录保留。

追加：`jsc-ts-contracts-02.trx` 36 通过，加入伪造编译角色/前端身份、修订变化和执行消息拒绝；`jsc-ts-actions.trx` 验证 TS actions/事件的单元格与虚拟文件原子提交和撤销。`jsc-compiler-budgets.trx` 3 通过，覆盖密集 1 MiB TS 完整检查，以及编译/执行两种启动策略的 OS 负向访问。

## 9 J4 桌面 TS 流程

Monaco 为每个模块建立 JS/TS 模型，两种语言共用精确逻辑 ID 的类型解析和 QuickExcel 声明，不含 DOM/Node。JSC 候选的语言按钮只允许 API 2 切换；实际应用与运行前将编辑器草稿 flush 到带版本的快照，隔离编译完整检查成功后才应用。失败保留原已应用源码和当前失败草稿，诊断绑定草稿版本，可直接定位；草稿变化后拒绝旧诊断。

编译进度显示在脚本窗，既有停止按钮及关闭窗口可以取消。保存允许保留尚有类型错误的原文，但不会因此执行或接受旧产物。编译期间文档/修订/源码改变会使返回结果失效。执行错误导航使用映射后的原始行列。

`artifacts/jsc/web-results/20261103-070909/result.json`：真实 WebView2/Monaco 窗口 48 项检查通过，约 16.16 秒；包括 TS 模型、失败编译保留草稿/旧源码、诊断定位、修正后应用运行、TS 原文保存/重开，以及原 JS、actions、事件/虚拟文件和既有 PDF 流程。不是物理输入、IME、辅助功能或干净机器验收。

## 10 J5 发布与独立平台

Windows 候选 `artifacts/avalonia/releases/20261103-072018-8458bc53/application`：发布定向检查 102 通过、1 项 `JSC-MEM-001` 明确延期；8 项发布包 self-check 通过，包含真正隔离的 TS 编译及运行。包包含 JSC、ICU 77.1、所需 VC Runtime、编译器 bundle、构建身份与对应源码/许可证；不包含 QuickJS、开发探针或故障夹具。默认发布指针未激活该候选。

`artifacts/avalonia/deployment/20261103-072326-40af7246/result.json`：普通用户中文路径、只读安装副本、缺少 Web 资源、缺少 Worker 四项通过。只读 ACL 副本不等于实际 Program Files 或干净机器。

Linux 使用独立 QEMU 写时复制磁盘 `artifacts/jsc/linux-vm/jsc.qcow2`，基线磁盘不变。Ubuntu 普通用户、clang 18、固定 ICU 77.1 源码、同一 WebKit commit、同一桥/编译器。首次 8 GiB WHPX VM 出现内核页错误/soft lockup，日志保留；改为 4 GiB 后 ICU 构建成功，Ruby 代码生成器一次异常退出，重试进入 C++ 编译。上述构建环境故障不是引擎验收，当前不据此标记 Linux 通过。编译 scope 单独采用 768 MiB/6 秒，其余 Landlock、seccomp、无 swap、32 tasks、父进程死亡、私有只读 payload 和租约规则不放宽。

系统 Ruby 3.2.3 后续又发生对象 ID/GC 崩溃，改为源码构建、固定哈希的 Ruby 4.0.9（带 libyaml/psych）作为仓库内构建工具。Ruby 生成阶段使用新构建的解释器自举；不进入产品载荷。失败日志及原配置继续保留在 `artifacts/jsc/linux-vm/`。

真实编译 Worker 性能采样发现 J0 原生探针未计入的托管/桥接内存：初始密集 1 MiB 样本触及 768 MiB Job，失败证据位于 `artifacts/jsc/worker-performance/windows-01` 至 `windows-05`。调整为先释放托管 bootstrap 临时对象，并为独立编译 VM 设置每轮分配 32 MiB 时请求 GC；这是软 GC 触发条件，**不实现或关闭 JSC-MEM-001**。调优后的定向测试和部分采样通过，但独立 `windows-08` 三次密集输入仍有一次在进程上限处失败，因此不能冻结为可靠的 768 MiB 编译配置。`windows-07` 与定向测试发生重叠，只保留原始数据，不作为正式性能结论。

J5 根据上述新证据将**独立编译进程**预算冻结为 1 GiB、策略 `typescript-check-1024m-5s-v1`，同时把编译策略加入产物缓存摘要；Windows Job 与 Linux scope/内核读回同步调整。执行仍为 256 MiB、256 KiB 栈、1 秒/5 秒；编译的 5 MiB 栈、5 秒/6 秒、无 swap/32 tasks、零能力访问隔离和源码/产物限额保持。J0 探针和 768 MiB 历史结果不回写为新配置通过。

`artifacts/jsc/worker-performance/windows-09/result.json`：独立顺序采样 24 项全部通过，记录完整 worker payload 摘要、源码摘要、分阶段墙钟、OS 峰值内存及 5 ms 线程采样。每种三次，未报告 p95。小 TS 冷编译含启动/IPC/清理中位 701 ms（681–757），密集 1 MiB 为 1516 ms（1501–1594），峰值提交分别约 331/784 MiB；满足本机小输入 ≤1 秒、密集输入 ≤2 秒目标。JS 小轮、64 模块、1 MiB 源码、256 次批调用及计算密集批次中位分别 299/262/286/280/289 ms。128 KiB 虚拟文件 actions 为 357–466 ms、事件为 373–390 ms，取消为 33–51 ms，编译缓存命中为 0.18–0.42 ms。这是本机进程级样本，不代表其他设备性能。

桌面最终复测 `artifacts/jsc/web-results/20261103-080415/result.json` 48 项通过，包含最新草稿缓存与诊断行列处理。

J5 最终 Windows 候选 `artifacts/avalonia/releases/20261103-081744-de4bda74/application`：110 项定向检查通过、1 项明确延期；8 项包自检通过。归档 SHA-256 `FE0837228B8456ACD2DD9000552EB1ECC96BE15CB84A7D746C1731219112EA7B`。部署四项复验在 `artifacts/avalonia/deployment/20261103-083044-0c1ef43e/result.json`，全部通过。

Linux 最终在独立 Ubuntu VM 普通用户下完成 J5：固定 JSC/ICU 原生构建成功，系统 Ruby 崩溃用固定来源 Ruby 4.0.9 构建工具解决。首次空 LocalApplicationData 导致的私有租约根问题改为先取得未验证绝对位置，再执行既有私有创建与 reparse 拒绝；未放宽文件访问或回收规则。编译角色 CLR 堆为 256 MiB，执行角色仍为 64 MiB，分别包含在 1 GiB/256 MiB 进程预算内。两角色内核保护和 32 项负向访问检查相同；没有关闭 AppArmor 或 WebKit 沙箱。

| Linux J5 证据 | 结果 |
| --- | --- |
| `artifacts/jsc/linux-vm/evidence/linux-smoke-runs/run-03/functional-02/result.json` | 8 项通过，含 JS/TS、actions、事件与虚拟文件 |
| 同目录 `sandbox/result.json` | 执行/编译两角色共 8 组通过；内存超限 SIGKILL，墙钟终止约 5.14/6.10 秒，32 tasks 与负向访问限制生效 |
| 同目录 `recovery/result.json` | 两角色宿主死亡后 worker 消失约 5.72/5.43 ms；日志化租约恢复通过，不清除未登记哨兵 |
| 同目录 `performance/result.json` | 独立顺序 24 样本通过，保留身份/分阶段时间/内存/线程原始数据 |
| `artifacts/avalonia/platform-releases/linux-x64-20261103-082630-aa2a62f8/` | 完整发布包；VM 内 Xvfb 自检 8 项通过，日志 `artifacts/jsc/linux-vm/release-self-check-01.log` |

Linux 小 TS 冷编译中位 2725 ms（2642–2726），密集 1 MiB 中位 3598 ms（3533–3624）；JS 小轮、64 模块、1 MiB 源码、256 次批调用、计算密集批次中位分别 2149/2169/2256/2233/2286 ms。VM 私有 payload 复制/fsync 占主要启动时间，未超过固定轮/编译截止；这些样本不代表满足 Windows 本机交互目标，也不代表 p95。原生失败构建日志与上述结果完整保留在 `artifacts/jsc/linux-vm/linux-evidence.tar.gz`，未覆盖失败证据。

Linux 发布、WebView 与输出在 Xvfb 中验证，是工程证据；物理 IME、DPI、辅助功能、真实打印和干净设备验收仍延期，Linux 中文 IME 无候选窗的问题未关闭。macOS 执行仍禁用，J6 不发布其引擎/worker。

## 11 J6 单引擎切换与回退

源码已去除全部 `QUICKEXCEL_JSC` 双路径条件，`JscSession` 只调用固定 JSC；原 QuickJS C 桥、CMake 目标和构建脚本已移除。进程内 `JscRuntimeFactory` 移到 Core.Test，生产组合根只保留隔离 worker。Windows/Linux 默认发布只接受 JSC；显式请求旧引擎构建直接报错。来源、历史 MIT 许可证、旧证据及 WPF 归档/tag 保留。生产原生构建只编译 JavaScriptCore 目标，开发探针需显式选择且发布拒绝携带。

切换前已逐项验证旧 QuickJS 包的 717 个文件摘要，完整复制到 `artifacts/avalonia/rollback/quickjs-20261103-040216-fb1af252/application`，旁边保留原 active 指针和回退身份。原 release.json SHA-256 为 `09EAE26F8C48F9FDC155733D90B320559FD62654E49200670623ABC399ABDB75`。

关闭应用后，可运行：

```powershell
.\scripts\select-quickexcel-package.ps1 -PackageDirectory .\artifacts\avalonia\rollback\quickjs-20261103-040216-fb1af252\application
```

选择器验证完整包摘要和已通过的包自检后原子更新指针；宿主/worker 必须整包回退。旧版会拒绝 `typescript-sources`，不能用其覆盖含 TS 的文件；升级前保留原文件副本。最终默认 JSC 包及复验结果如下。

### 最终发布与验收

| 项目 | 本轮最终证据 |
| --- | --- |
| 默认构建 | `artifacts/jsc/final-solution-build.log`，solution Release 0 警告 / 0 错误；最后增量复核 `final-build-incremental.log` 同样通过，无重复原生构建 |
| 旧 JS / 模块 / 批次 / 动作 / 事件 / 格式 / 签名 | `artifacts/jsc/test-results/j6-script-contracts.trx`：311 项通过；只运行相关区域，未运行全仓回归 |
| 原生 smoke | `artifacts/jsc/final-native-smoke.log`：通过；栈溢出/不可捕获中断使上下文失效，不复用；明确报告开放内存 issue |
| 真实桌面 | `artifacts/jsc/web-results/j6-final/result.json`：48 项通过，约 16.91 秒；TS 全流程和原 JS/API v1 保留 |
| Windows 发布检查 | `artifacts/avalonia/releases/20261103-084654-7b4f8355/evidence/release-contracts.trx`：110 通过，1 项 `JSC-MEM-001` 显式延期；隔离执行/编译、取消、故障与恢复在内 |
| Windows 发布/启动器 | 同目录 `self-check/result.json`：8 项通过；`artifacts/jsc/launcher-final-self-check/result.json`：默认启动器实际启动最终包，8 项通过 |
| Windows 部署 | `artifacts/avalonia/deployment/20261103-084851-fb45346e/result.json`：普通用户中文路径、只读副本、缺少 Web、缺少 Worker 四项通过 |
| Linux 最终包 | `artifacts/avalonia/platform-releases/linux-x64-20261103-084841-54e94e61/application`；普通用户 VM / Xvfb 自检 8 项通过，结果 `artifacts/jsc/linux-vm/final-release-self-check-result.json` |
| macOS 边界 | `artifacts/avalonia/platform-releases/osx-arm64-20261103-085023-63f7d3d7/application` 交叉发布成功；516 个文件，无 JSC/QuickJS 原生引擎、compiler bundle 或 ScriptWorker，脚本身份 `disabled`；不是 macOS 实机验收 |
| 回退 | `artifacts/jsc/rollback-check/result-02.log`：旧完整包能打开纯 JS 控制文件，以“包含未知必需能力”拒绝 TS 文件；`select.log` 验证旧包选择器，之后默认构建重新激活最终 JSC 包 |
| 单引擎门禁 | 源码无 QuickJS 类/生产引用/原生目标/复制规则；`artifacts/jsc/removed-engine-build-rejection.log` 证明显式旧引擎参数被拒绝；发布脚本拒绝旧引擎、开发探针和测试夹具 |

Windows 当前默认目录：`artifacts/avalonia/releases/20261103-084654-7b4f8355/application`，版本 `0.6.0-preview.8`。`artifacts/avalonia/active-windows.json` 已原子更新，736 个文件摘要及 active manifest 再次逐项校验通过，记录在 `artifacts/jsc/final-active-integrity.json`。

- Windows release.json SHA-256：`82B6DA87276A6D0F8FECD4BE6B94F9FBC16399DCB1E5FF0A52D86680932294EF`。
- Windows ZIP：`artifacts/avalonia/releases/QuickExcel-Avalonia-0.6.0-preview.8-win-x64-20261103-084654-7b4f8355.zip`，SHA-256 `8D77865341BA4FAD91535AC6116A1D53DA2B1B533EA4899987E7AE9AEC8ABDA0`。
- Linux 归档：最终 Linux 包目录的上一级 `QuickExcel-linux-x64.tar.gz`，SHA-256 `14476AC7B6CBA6B4CA09563327F947832BED4844A04068AD342837E326EFC599`。

从旧包回到本轮最终版时，关闭应用并运行：

```powershell
.\scripts\select-quickexcel-package.ps1 -PackageDirectory .\artifacts\avalonia\releases\20261103-084654-7b4f8355\application
```

最终构建时发现 ICU 方案的并行目标共享中间目录，出现 tlog/data 文件争用；改为串行构建 ICU，JSC 保留并行构建。失败日志 `artifacts/jsc/logs/20261103-084100-7e8eaab5/` 保留；修正后的 `20261103-084220-fb4f2fe8/` 通过。原生构建以成功身份文件作为增量标记，缺失原生库时重新构建，避免 Ninja 未改写 DLL 时间戳引发每次托管构建重复准备。编译器准备使用 Windows `python` / Unix `python3`，校验固定 TS 归档及已解压缓存。

回退验证的首个测试驱动错误地要求旧异常消息含新特性名；旧版实际明确报告未知必需能力，未打开或丢失 TS。修正驱动并加入旧版可打开纯 JS 的对照后通过，初始日志保留。未修改旧包来制造兼容结果。

**完成边界：** 70 号文档的 J0–J6 按用户调整后的门禁完成；`JSC-MEM-001` 仍为 **OPEN / deferred / non-blocking**。其余既有 M4 严格策略、M6、安全/部署、真实 IME/DPI/辅助功能/打印/干净设备门禁保持独立；未因引擎迁移自动关闭。Linux VM 性能与 Windows 本机数据按各自条件报告，macOS 脚本继续禁用。所有失败记录、历史来源/许可证、旧完整回退包和 WPF 归档/tag 保留。

## 12 脚本编辑辅助与 API 别名

2026-11-03。编辑器现提供实时 TypeScript 语法/语义错误标记、当前模块错误/警告计数、API 和变量补全、悬浮类型与文档、函数参数提示。工作区的“JS / TS”切换模块语言；JavaScript 保留原有 JSDoc 检查，含 TS 的模块集沿用固定的严格编译选项。底部提供“补全”“参数提示”“类型提示”“下一问题”按钮，保留当前光标位置。

| Windows 快捷键 | 功能 |
| --- | --- |
| `Ctrl+Space`，或输入成员访问的 `.` | 补全列表；可接受所选项写入草稿 |
| `Ctrl+Shift+Space` | 函数参数提示 |
| 悬停，或 `Ctrl+K` 后 `Ctrl+I` | 类型与 API 文档 |
| `F8` / `Shift+F8` | 当前模块下一处 / 上一处问题 |

此前 `editor.api` 只安装基础编辑器，语言 worker 已能返回诊断与补全，却没有对应的界面组件。本次显式注册 suggest controller、hover、parameter hints、goto error 和图标。Monaco 固定为 0.57.0；该版 `features/suggest/register` 只注册 inline suggestions，因此补全框直接引入随包的 controller。升级必须复验实际弹窗和插入行为；不以 worker 返回候选词代替 UI 验证。参考 [Monaco ESM 集成](https://github.com/microsoft/monaco-editor/blob/main/docs/integrate-esm.md)与[编辑器选项](https://microsoft.github.io/monaco-editor/typedoc/interfaces/editor_editor_api.editor.IEditorOptions.html)。

语言服务仍只同步本工作簿的逻辑模块与 ES2022 / QuickExcel 声明，没有 DOM/Node 类型。诊断仅针对附着于编辑器的当前模块，并在切回模块时重新检查，使依赖模块草稿变化生效；模型继续保留独立撤销历史。状态栏计数不是全包检查结果，执行前仍由隔离编译 worker 完整检查。编辑辅助不执行脚本、不逐键提交工作簿；只读禁用补全插入按钮，组合输入期间禁用辅助按钮及 F8 跳转。

`$` 是 `quickExcel` 的同一冻结对象别名，JS API v1、API v2 模块和 TypeScript 均可使用：

```typescript
$.functions.register('DOUBLE', (value: number) => value * 2);
$.actions.register('写入', ({snapshot, patch}) => {
  patch.setValue(snapshot.sheets[0].id, 'A1', 42);
});
$.events.on('cell.click', ({event}) => console.log(event.name));
```

动作与事件仍要求 API v2 的 actions 绑定及用户授权；函数、动作、事件使用原注册表、执行隔离与原子补丁规则。全局 `$` 不可重写/配置，原 `quickExcel` 写法继续有效，模块或 v1 函数体内已有的局部 `$` 可以正常遮蔽别名。声明使用 `declare const $: typeof quickExcel`，供 Monaco 与隔离编译器共用，不复制两套 API 类型。

别名兼容性测试发现预解析源码没有使用编译器的 `moduleDetection: Force` 回调，无 import/export 的 TS 文件误成为全局脚本。内存 compiler host 现将 `getSetExternalModuleIndicator(options)` 传入 `createSourceFile`，使本来就约定为 ESM 的源码保持模块局部作用域，避免局部 `$` 与全局声明冲突，也拒绝未导入的其他模块局部变量。未修改源文本，源码映射和诊断坐标继续按原文生成；编译器身份随 frontend/声明自动更新。语义依据：[TypeScript moduleDetection](https://www.typescriptlang.org/tsconfig/#moduleDetection)。

验证范围：

- `scripts/test-avalonia-web.ps1 -ScriptsOnly`：真实 Windows WebView2 **56 项通过**，`artifacts/avalonia/web-tests/20261103-103057-d08305d3/result.json`。覆盖实际补全显示/插入、JS/TS `$`、错误与修正、参数/悬浮文档、跨模块类型、DOM/Node 缺失、草稿/保存/恢复、只读、64 模块以及隔离编译/应用。新增 `-ScriptsOnly` 跳过无关 PDF 检查。
- Core 定向 46 项：编译器、别名与草稿契约首轮 **44 项通过**，两个新注册限制测试因断言的异常类型与现有 worker 行为不同失败；修正断言后这 **2 项通过**。证据分别为 `artifacts/script-editor-tests/script-editor-final.trx`、`script-alias-guards.trx`。含 JS v1/v2 与 TS 对象身份、局部变量遮蔽、注册限制、动作/事件原子提交和撤销、类型错误拒绝、模块作用域、取消与源映射。不运行全仓回归或性能基准。
- 首轮 UI 检查暴露精简 suggest 注册入口不安装控制器；后续夹具修正了合成输入事件来源和 LF/CRLF 光标偏移。失败证据保留于同一 `artifacts/avalonia/web-tests/` 下，未改动旧发布包。测试辅助分支仍由正式 Web 构建裁去。

这些是自动工程检查，不新增物理 IME、辅助技术或独立 Linux/macOS 实机结论。默认发布指针与 §11 历史包未因本次源码更新而切换；重新发布需使用仓库既有完整包流程。
