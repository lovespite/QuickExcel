# 工作簿脚本数字签名与证书信任

2026-10-29。此功能更新此前文档中“打开一律不运行脚本”的桌面策略；未签名文件和恢复副本仍保持停用。存储层与 `WorkbookDocument.Open` 只读取文件，只有桌面完成签名验证与信任判断后才启用脚本。AppContainer、Job、IPC 和动作事务边界保持不变。

## 使用

在“脚本”页面左侧点击“签名…”，选择带私钥的 PFX/P12 代码签名证书，输入密码；也可“创建自签证书并签名…”，设置证书名称、私钥文件保存位置和至少 8 字符的 PFX 密码。新建证书为 RSA 3072，具有数字签名和代码签名用途，有效期两年。签名完成后保存 `.xlsq`。

| 状态 | 打开文件时的行为 |
| --- | --- |
| 未签名 | 保持停用；仍可点击“运行”手动授权 |
| 签名有效，系统证书链验证通过或已信任自签证书，且已确认当前引擎策略 | 自动启用、重算并发现操作脚本 |
| 签名有效且证书受信任，但尚未确认当前引擎策略 | 保持停用；显式“应用并重算”后记住当前策略 |
| 签名有效，证书为未信任的自签证书 | 显示主体、颁发者、有效期、SHA-256 指纹，默认“取消” |
| 选择“仅本次允许” | 仅当前会话允许执行，不保存信任 |
| 选择“始终信任此证书”并成功应用 | 记住完整证书的 SHA-256 指纹和当前策略确认，以后此证书签署的工作簿在相同策略下自动启用 |
| 内容或签名篡改、证书失效、用途不符、链失败 | 保持停用，不提供信任无效签名的选项 |

操作脚本只发现并列出，动作体仍需点击“执行动作”。模块初始化和 UDF 会在启用及重算时执行。信任证书意味着允许它签署的其他工作簿执行脚本，自签证书显示的名称不证明签署者身份。

“签名详情…”可查看验证结果、移除签名；“受信任证书…”可移除记住的自签证书，同时停用当前工作簿脚本。Windows 系统信任通过系统证书管理维护，QuickExcel 不向 Windows 根证书存储写入证书。

## 覆盖与失效

2026-11-04 修复：之前 JSC 迁移提示在每次打开时固定显示，显式应用没有持久化策略确认。现在按证书指纹记录引擎、执行策略、IPC 版本、编译器身份、编译策略和选项组成的摘要，与证书信任一起使用原有按用户/路径隔离的受保护存储。旧版受保护指纹列表仍可验证证书，但不隐式授予新策略授权；成功显式应用后再打开会自动启用。策略变化需再次确认，“仅本次允许”不保存策略确认，移除证书同时移除对应确认。每次打开仍重新验证签名和证书，确认记录不能覆盖失效或篡改。打开时仅发现操作，操作体仍需用户执行；已授权的事件按原有事件契约分派。

签名专项测试覆盖确认持久化、其他工作簿/签署者、策略变更、旧受保护列表、撤销和篡改。Avalonia `scripts/test-avalonia-web.ps1 -ScriptsOnly` 覆盖真实应用与重开路径，以及自动授权不运行操作体。

签名覆盖所有模块 ID、路径、原始源码（含库模块）、API 版本、入口顺序、全部 workbook 用途绑定，以及签名版本、算法和证书 SHA-256。模块和绑定 ordinal 排序，入口保持原序，源码换行及空白保持原样。

单元格数据、公式、工作表、工作簿身份、修订号和缓存不在签名范围内；数据编辑、重算、保存和另存保留签名。签名不能证明数据可信或脚本无害。

修改源码、模块或用途会移除签名并撤权，无效编辑器草稿也立即移除旧签名。单元格撤销/重做不恢复旧签名。签名变更经 `EditorController` → `DocumentSession` 提交，标记 dirty，不进入单元格撤销历史；提交要求模块仍与签名时捕获的对象相同。

## 文件与算法

### 签名时编译与 XLSQ 编译缓存（3.4.0–3.6.0，2026-11-05）

“签署全部脚本”先在既有零能力编译 worker 内完整检查并生成 JS 和 Source Map，成功后才签名；不会执行初始化、UDF、动作或事件。JavaScript 和 TypeScript、API v3/v4 都适用。失败显示模块和行列诊断，不生成新签名；保存工作簿时将缓存一起写入 XLSQ。

必需能力为 `script-compilation-cache`。3.5.0 改为缓存格式 v2，清单分别以 `scriptCompilation` 和 `scriptCompilationData` 引用两个独立 ZIP 条目，各自带 SHA-256：

- `cache/script-compilation.json`：最多 64 KiB，仅保存完整源码签名摘要、编译器/源/输出身份、数据条目的长度和 SHA-256、各模块 ID 与字节偏移/长度、保护方式及证明；不嵌入 JS 或 Source Map 正文，也不把二进制数据 Base64 塞回 JSON。
- `cache/script-compilation.bin`：按模块索引顺序连续保存 JS、Source Map 的原始 UTF-8 字节；总计最多 5 MiB，分别遵守 1 MiB JS 和 4 MiB 映射限额。读取先验证元数据、偏移边界、总长度、二进制摘要和证明，之后才解码对应文本及验证映射。

原始源码、语言和固定选项仍为文件的权威内容；缓存不是 JSC 字节码或执行上下文。3.4.0 的格式 v1 单体 JSON 缓存作为失效缓存丢弃，不反序列化其编译正文；执行前重编译，保存时写出 v2 两条目。旧 JSON 仍受原 16 MiB ZIP 展开上限约束，但只有不超过 64 KiB 的 v2 元数据才允许进入缓存 JSON 解析流程。

3.6.0 在“设置 → 脚本与 API → Source Map（源码映射）”提供本机编译选项，中英文名称与说明随界面语言切换：

| 档位 | English | 编译及诊断行为 |
| --- | --- | --- |
| 无 | None | `sourceMap: false`，不生成映射，缓存中的映射长度为零；运行错误标注为生成 JS 的位置 |
| 精简（默认） | Compact | `sourceMap: true, inlineSources: false`，保留行列映射，不含 `sourcesContent`；用 XLSQ 已有原文定位错误 |
| 完整 | Full | `sourceMap: true, inlineSources: true`，保留行列映射和嵌入的原始源码 |

选项保存到本机 `settings.json` 的 `ScriptSourceMap`，旧设置文件未指定时使用精简。关闭设置窗口后生效：停用当前脚本及其持久上下文、失效当前编译缓存；下次显式应用或签名按新档位编译，不自动重放初始化。选项进入编译请求摘要与缓存元数据/证明，防止三档缓存混用；源码签名格式和有效性不因调试产物选项变化而改变。XLSQ 元数据记录产物实际档位用于验证，打开文件不会覆盖本机偏好。

3.6.0 验证：Release 构建 0 警告 / 0 错误；53 项定向检查通过，`artifacts/source-map-options-tests/source-map-options.trx`，包含三档实际编译/错误位置、本机设置保存与旧设置默认、模式切换后的缓存及签名保存重开，并明确排除 1 MiB 密集输入压测。实际中文、英文 Settings/签名入口各 29 项通过，分别位于 `artifacts/avalonia/workspace/20261105-135750-cd754159/result.json` 和 `artifacts/source-map-options-ui/english-f86b0f2949674e89a33b48d9d4769d24/result.json`；所有用户设置均隔离到测试目录。

3.6.0 正式本地包 `artifacts/avalonia/releases/20261105-135917-b4bb592e/application` 经 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 发布，28 项生产自检通过；active 指针、release.json 摘要和启动 EXE 摘要已核对，默认启动器使用此包。

签名时另以同一发布者证书签署缓存摘要，不修改原有纯 JS v1 / TS v2 源码签名格式。普通 ZIP hash 不能防止攻击者同时改写 JS 与 hash；缓存证明也必须验证。打开只读取数据；在既有证书信任和策略授权通过、准备执行时，源码签名、缓存格式、编译身份、摘要、证明及映射全部匹配则直接使用缓存，不启动编译 worker。

缺少缓存、缓存版本/源码签名/编译身份不符、缓存 hash/证明/JSON/映射损坏时，从已签名原文重新编译，成功替换会话缓存后才执行，下次保存写入新缓存。无需发布者私钥：重建产物使用 Windows DPAPI 当前用户或 Linux Secret Service 的受保护本机凭据认证，不能改变源码签名或授权。此类重建缓存换设备/用户后会重新编译；重新签名可生成可随文件携带的发布者证明。本机保护不可用时仅当前会话复用，下次打开重新编译。源码签名本身无效仍拒绝执行。

编辑器草稿一变化便使缓存失效并撤销执行授权；应用源码、语言、模块或用途变更会移除签名，保存时删除旧缓存的两个条目。缓存派生状态按签名对象隔离，不进入 Core 领域对象、单元格历史或虚拟文件快照；脚本不能通过 `//xlsq/cache/script-compilation.json` 或 `.bin` 写入系统缓存，避免绕过验证或占用虚拟文件/桥接容量，普通 `cache/` 虚拟目录及其他用户文件仍可使用。缓存重建不隐式保存用户文件，保持原有保存冲突检查和原子替换流程。

3.5.0 按用户要求不开展大缓存压测，也不重跑相关回归；同步修改既有测试夹具以适配二进制条目，仅执行构建与仓库规定的发布自检。解决方案 Release 构建 0 警告 / 0 错误，正式包 `artifacts/avalonia/releases/20261105-134542-857b9f2e/application` 的 28 项生产自检通过，已激活并核对 active 指针、release.json 摘要和启动 EXE 摘要。

以下为 **3.4.0 原缓存格式的历史验证**，不能代替 3.5.0 验证。

本轮验证：144 项相关 Core 测试通过（`artifacts/compilation-cache-tests/compilation-contracts.trx`、`cache-drafts-virtual-files.trx`；虚拟目录边界修改后对应 1 项复验见 `cache-virtual-boundary.trx`）。缓存测试覆盖 JS/TS 与 API v3/v4、重开时编译进程启动数为 0、10 类缓存失配/损坏后的重建及再次保存重开、草稿/源码/语言/绑定失效、签名失败/取消、签名时不执行初始化、UDF/Actions 共用产物与虚拟文件边界。首轮夹具将 API v4 UDF 名写成未限定的 `VALUE`，修正为既有 `JS.main.VALUE` 后通过；失败记录 `compilation-cache.trx` 保留。

真实 Avalonia 签名入口 `scripts/test-avalonia-workspace.ps1 -Mode Commands` 24 项通过，证据 `artifacts/avalonia/workspace/20261105-133010-438b23ff/result.json`。按仓库要求执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`：正式 Windows 包 `artifacts/avalonia/releases/20261105-133201-c88b2dc0/application`，版本 3.4.0，生产自检 28 项通过并激活；未运行全仓回归，也未新增 Linux/macOS 实机结论。

XLSQ 保持格式 v2，新增必需能力 `script-signatures`，清单 `scriptSignature` 指向 `security/script-signature.json` 并带条目 SHA-256。旧读取器拒绝此能力；移除签名后保存删除能力、引用和条目。签名 JSON 上限 64 KiB，复用 ZIP/JSON 累计预算和重复属性检查。

JSON 为 `version: 1`、`algorithm`、Base64 DER `certificate` 和 Base64 `value`。支持 RSA-PSS-SHA256（至少 2048 位）与 ECDSA-P256-SHA256（P1363 固定长度）。私钥/密码不入包，PFX 私钥以 `EphemeralKeySet` 导入。

载荷用 .NET `BinaryWriter` 写入：严格 UTF-8，字符串长度为 7-bit 编码字节长度，Int32 小端。顺序为 `QuickExcel.ScriptSignature` 字符串、版本 Int32（1）、算法字符串、大写证书 SHA-256 字符串、API Int32、模块数量及排序后 `(ID,path,source)`、入口数量及入口序列、绑定数量及排序后 `(module,"workbook",kind)`。普通包摘要不能代替数字签名。

## 信任策略与限制

每次打开和点击“运行”重新验证。证书必须有代码签名 EKU，KeyUsage（若存在）允许 DigitalSignature，当前时间在有效期内。系统链使用代码签名策略、Online 撤销检查（排除根）、5 秒 URL 检索超时；撤销未知/离线等失败不自动信任。参考 [.NET X509ChainPolicy](https://learn.microsoft.com/dotnet/api/System.Security.Cryptography.X509Certificates.X509ChainPolicy)。

只有链仅含该证书、subject/issuer 一致且唯一错误为 UntrustedRoot 的自签证书可以被记住。指纹不覆盖过期、撤销、显式不信任、用途或其他链错误。同名、换密钥或重新签发均不是原证书。无时间戳服务，过期后需重新签名；自签证书无在线撤销服务，应移除不再信任的证书。

信任列表 `%LOCALAPPDATA%/QuickExcel/trusted-script-certificates.json` 使用版本化 JSON 外壳，内部指纹、显示名称和到期时间由 Windows DPAPI `CurrentUser` 保护并验证完整性；用途和列表绝对路径参与绑定，其他路径的受保护列表不能直接替换。跨进程独占文件锁和原子替换保护写入。修改保护数据、追加未保护字段、格式错误、解密失败或写入失败均不授予授权；每次验证重新读取，不能沿用上一次授权结果。参见 [微软 DPAPI 说明](https://learn.microsoft.com/dotnet/api/system.security.cryptography.protecteddata)。

旧版明文列表无法证明用户曾同意，升级后拒绝使用，也不会在新的信任写入中自动迁移。遇到旧版或损坏列表时脚本保持停用；关闭程序，将上述文件移走或删除，重开工作簿后重新核对证书并确认。删除列表只会撤销记忆的信任，不会使自签证书自动受信任。测试及 `--ui-validation` 使用指定目录中的独立列表。

此保护针对直接修改配置内容，不隔离已经以同一 Windows 用户身份执行的恶意代码：该身份也能调用 DPAPI。它也不提供同一路径历史完整密文的防回滚保证。CA 吊销状态仍由每次证书链验证决定，保存的指纹不能覆盖吊销结果；自签证书的移除信任不等同于 CA 的在线吊销。

文件不保存“已信任”标志。恢复入口不自动执行，用户审阅后可手动运行。`.xlsx` 导出不含脚本及签名。此功能不等同于程序/安装器 Authenticode 签名或完成 M6 OS 门禁。

## 验证

按修改范围运行，未执行全量回归：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --no-build --artifacts-path .\artifacts\build -v minimal --filter "FullyQualifiedName~ScriptSignatureTests|FullyQualifiedName~TypedStorageTests|FullyQualifiedName~DocumentSessionTests|FullyQualifiedName~M3StorageTests|FullyQualifiedName~WorkbookFacadeTests"
.\scripts\test-script-signatures.ps1
```

Release 构建 0 警告/0 错误。正常用户环境相关测试 125/125（签名 26 项），结果 `artifacts/validation/script-signatures/related-tests-final.trx`。受限环境首轮 105/106，PFX 导入因环境权限失败；正常用户测试通过，未更改生产权限要求。

WPF 专项覆盖真实信任按钮、默认取消、一次/持久信任、重开自动授权、动作仅发现、数据编辑保留签名、源码编辑撤权、篡改拒绝、损坏信任列表撤权、恢复入口和撤销信任。公式和动作发现使用实际 AppContainer worker。截图、逐项结果和二进制哈希位于测试脚本输出的唯一 artifacts 目录；程序化断言不代表真实鼠标/键盘或完整 OS 验收。

### 信任列表防篡改修复（2026-10-29）

已用伪造明文列表回归测试复现旧实现接受注入指纹的问题，修复前证据为 `artifacts/validation/script-trust-hardening/trust-tamper-before.trx`。修复后签名专项新增伪造列表、替换合法列表、修改保护数据、追加明文指纹、跨路径替换、删除后重新确认，以及真实 CRL 吊销测试。吊销夹具使用独立 CA 和唯一回环 HTTP CRL，不安装系统根证书；先断言未吊销证书可通过，再断言吊销证书返回 `Revoked`，有无匹配指纹都不能自动启用。

本轮 Release 构建 0 警告/0 错误，仅运行 `FullyQualifiedName~ScriptSignatureTests` 和 `scripts/test-script-signatures.ps1`，不运行全量回归。签名专项 34/34 通过，证据为 `artifacts/validation/script-trust-hardening/signature-security-final.trx`。DPAPI/PFX 和回环 CRL 使用正常 Windows 用户环境验证；受限令牌下的用户配置保护失败不通过降低安全策略处理。WPF 新增伪造列表阻止打开/运行、保护数据篡改撤权和重开拒绝，共 23 项断言通过，证据位于 `artifacts/validation/script-signatures/20261029-022455-eed7142b`。
