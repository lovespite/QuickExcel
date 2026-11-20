# 脚本 fetch 与网络面板

2026-11-20，产品 6.1.1。文件脚本 API 仍为 v4（保留 v3 兼容执行），IPC 为 13，原生桥标识仍为 `qx-async-3`。6.0.0 引入 fetch 与异步 Actions，6.1.0 增加本机网络开关和可配置限额，6.1.1 修复同一会话重复加载脚本时的重复网络授权。此文补充 80 号入口约定，并替代其中初始化 / Actions 必须同步的限制；UDF、Events、控制台仍遵守原有执行约定。

## 声明和授权

在初始化入口 `main.js` 或 `main.ts` 的顶层声明完整域名或 IP：

```js
host.permissions.network.request("api.example.com");
host.permissions.network.request("127.0.0.1");
host.permissions.network.request("::1");

export function main(workbook, host) {
  // 可选的工作簿初始化；这里不能 fetch。
}
```

不接受 URL、路径、端口、通配符或子域名通配。域名按 IDN ASCII、大小写及尾点规范化；IPv4 / IPv6 使用规范地址。授权某个域名不授权其子域名或其他重定向目标。端口由请求 URL 指定。声明最多 64 项，只允许在初始化根模块顶层完成；导入的辅助模块、Actions、Events 和运行期间不能追加声明。拼写仅为 `request`。

宿主首次加载入口后、执行初始化前显示完整地址清单，说明 Actions 可向这些地址发送工作簿数据。用户明确同意后方可执行。授权只属于当前打开的工作簿会话；重复加载、修改源码 / 绑定后重新应用时，规范化、去重后的完整地址清单不变即复用授权，不再弹窗。清单增删或替换会先丢弃旧授权，再为新的非空清单请求同意；空清单无需弹窗，不累积历史清单的权限。明确停用脚本、关闭网络访问或关闭会话仍撤销授权。重建运行上下文照常取消旧请求、废弃旧授权弹窗结果并重新执行显式初始化，不保留旧 worker。拒绝授权不会执行初始化，也不会发送请求。用户阅读弹窗期间没有活动 worker 轮次计时。

## 本机设置（6.1.0）

设置 → 脚本与 API → **脚本网络与诊断限额**保存本机用户偏好，不写入工作簿，也不能由工作簿脚本修改。

| 设置 | 默认值 | 可选范围 |
| --- | --- | --- |
| 允许访问网络 | 开启 | 开启 / 关闭 |
| 单次请求 / 响应大小上限 | 各 8 MiB | 1–10 MiB 整数 |
| Action 执行超时（含网络等待） | 10 秒 | 1–300 秒整数 |
| 网络面板内存上限 | 50 MiB | 1–512 MiB 整数 |
| 网络面板记录条数上限 | 200 条 | 1–5000 条整数 |

关闭“允许访问网络”立即拒绝新的网络权限申请和请求，取消正在等待的 HTTP 操作，并撤销已有网络授权；不弹出可覆盖该开关的授权提示。重新开启不会恢复旧授权，访问前须再次取得用户同意。取消中的 Action 不提交工作簿部分修改。

点击“应用网络限额”保存四项数值。正文和执行超时使用下一次 Action 开始时的设置；调低面板限额会立即按时间释放旧正文或淘汰旧记录。Action 超时覆盖该 Action 内的脚本执行和网络等待，不包含用户阅读授权弹窗的时间。

## Actions 示例

```js
export async function main(workbook, host) {
  const response = await fetch("https://api.example.com/items", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ query: workbook.getActiveCell().getValue() })
  });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const result = await response.json();
  workbook.getActiveCell().setValue(result.name);
}
```

也可以从普通 `main` 返回 Promise 链。宿主等待 Promise 完成，沿用现有旋转遮罩和 Esc 取消；完成后作为一个历史事务提交。Promise 拒绝、超时、取消、过期结果、权限撤销均不提交工作簿 / 虚拟文件 / storage 的部分修改。HTTP 服务已收到的请求属于外部副作用，不能随工作簿撤销。

初始化 `main` 支持 async / 返回 Promise，但不能 fetch。UDF、Events、控制台、模块顶层均不能 fetch。Action 必须等待或返回其发起的全部 fetch；脱离 Action 生命周期的请求不会被保留。没有计时器或其他后台事件循环；永不完成且没有待处理 fetch 的 Promise 会失败。

## Fetch 行为及边界

实现参照 [MDN 使用 Fetch](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API/Using_Fetch) 的请求 / 响应模型，由 QuickExcel 自有代码实现，没有引入第三方 Fetch polyfill。

- `fetch` 返回 Promise；HTTP 404 / 500 正常返回 Response，`ok` 为 false。网络错误、未授权地址、非法参数以 TypeError 拒绝；非 Actions 调用直接报错。
- 提供 `Headers`、`Request`、`Response`，支持请求方法、请求头、字符串 / ArrayBuffer / TypedArray / Blob / URLSearchParams / FormData 请求体，以及 `text()`、`json()`、`arrayBuffer()`、`blob()`、URL 编码 `formData()`、`clone()` 和 `bodyUsed`。
- Response 支持有界缓冲区的 `body.getReader()`；完整响应先在宿主接收并校验限额，再送入 worker，并非实时网络流。文本解码仅支持 UTF-8；multipart 响应解析、文件系统 File、浏览器 DOM、Service Worker 不提供。
- 请求仅接受绝对 HTTP(S) URL。`redirect` 支持 follow / error / manual；manual 返回状态 0 的 opaqueredirect。follow 最多 10 跳，每跳检查地址授权，禁止 HTTPS 降级，跨源去掉 Authorization，按 HTTP 规则转换 301 / 302 / 303 的方法及请求体。
- 不共享浏览器 Cookie、系统登录凭据、代理凭据或本地 Cookie jar；`credentials` 接受 omit / same-origin，include 明确拒绝。Host、Cookie、Origin、Content-Length 等受限制请求头不能由脚本指定，Set-Cookie 不暴露给脚本。
- 这是桌面宿主权限模型，没有网页 origin，默认 cors 仅兼容参数形式，不模拟浏览器 CORS / opaque 响应。no-cors / same-origin mode 不支持。没有缓存；default / no-cache / no-store / reload 都实际发出请求，其他缓存模式拒绝。
- `AbortController` 支持预取消及尚在队列中的请求取消。网络交换通过既有 IPC 串行处理，Promise.all 可组织多个请求但传输按队列顺序；等待宿主响应期间不执行脚本回调。Esc、会话撤销 / 关闭会取消正在进行的 HTTP 操作。

## 隔离和内存

worker 保持零网络能力 AppContainer / Seatbelt / Linux 沙箱；只有宿主的专用 HTTP broker 发送网络请求。宿主独立验证当前操作确为 Actions、已同意的入口声明、逐跳目标、方法、头、正文大小和取消令牌，不依赖 JS 自检。网络调用复用已绑定操作身份和序号的 ModelCall 管道；生产 JSON 使用显式写入或源生成元数据。

每个 Action 最多 16 次 fetch，单个请求 / 解压后的响应正文分别受用户正文设置约束，默认各 8 MiB；头部不超过 32 KiB / 128 项。IPC 13 将宿主选定的正文上限传入 worker，宿主和运行时分别验证；脚本不能通过修改全局变量扩大上限。

16 MiB 单帧及整轮累计 IPC 预算保持不变。fetch 的内部正文编码使用带填充的 base64url，避免普通 base64 中的 `+` 被 JSON 转义放大；正文设置最大 10 MiB 为帧和元数据留出空间。单个 8 MiB 正文可装入帧，但一个 Action 同时上传 8 MiB 并下载 8 MiB，或多次请求累计较大正文，仍可能先触及累计 IPC 预算。工作簿模型调用也计入同一预算。

大正文传输使用 worker 专用的 JSON 值回调及直接 model 帧写入，避免重复生成 UTF-16 字符串和中间 JSON 树；完整帧通过 16 MiB 校验后才写入管道。受限缓冲区的容量扩张也不能超过预算。大块转换在 VM 空闲时回收 CLR 临时副本，仍计入宿主整轮期限；这些优化不扩大 256 MiB 进程限制，也不代表实现了 JSC 引擎分配硬上限。

网络等待不占用每次原生 JS 的 1 秒看门狗；用户 Action 的整轮宿主期限默认为 10 秒，可按设置调整。UDF、初始化、Events 和控制台的 5 秒宿主整轮上限保持不变；每次原生执行 / 续接仍为 1 秒。每次原生续接均在看门狗内清空 Promise 微任务；无限循环和微任务循环仍会终止 worker。Linux 仍有独立的 5 秒 OS worker 生命周期，设置更长的 Action 时间不会扩大该限制；其一次性 worker 仍不支持 v4 持久初始化，不能据此宣称跨平台能力已相同。

worker 中 Response 正文保留预算独立于网络面板，取 `max(8 MiB, 2 × 单个正文上限)`，默认 16 MiB，并最多保留 64 份。最旧正文超限后释放，读取被释放的 Response 会明确失败；部分读取的 reader 继续计入预算，不能绕过淘汰。完整消费或取消后立即解除运行时保留。用户自行保存解析后的对象、字节数组或 Blob 仍计入既有 worker 内存限制，不能用响应清理替代引擎内存边界。`JSC-MEM-001` 仍 OPEN。

## 网络面板

脚本工作区底部新增“网络”标签。仅面板可见且选中时记录；切换标签、隐藏输出或关闭工作区即停止记录。列表显示请求名称 / 方法、状态、大小及总耗时，支持筛选和清空；详情提供标头、载荷、JSON 预览、响应文本及计时。内容只在当前文档内存中保存，不写入工作簿或磁盘。

默认最多 200 条记录、50 MiB 保留内存，均可在设置中调整。内存预算同时计入请求 / 响应正文以及 URL、请求头、状态等元数据的保守估算，不仅统计正文；这是面板保留数据的预算，不是整个应用工作集的硬上限。超限按时间释放旧正文，记录仍显示“内容已释放”；若元数据仍超限则继续淘汰旧行。记录数超限、清空和会话关闭也释放正文。即使旧行被选中，保留正文引用也会置空；详情仅保留当前标签的最多 128 KiB 文本预览。此面板统计总耗时，不声称提供浏览器 DNS / TLS 分阶段测量。

## 验证

相关测试覆盖 broker 的本机 HTTP 服务、规范地址、禁止头部、跨地址重定向、HTTP 错误、压缩后正文限制、取消；原生 worker 集成覆盖授权顺序、async 初始化 / Action、返回 Promise、超过原生 1 秒的网络等待、拒绝后无请求、非 Actions 拒绝和事务回滚；Core 覆盖授权失效与记录淘汰。

### 6.1.1 会话授权复用验证

`ScriptNetworkSessionTests`、`ScriptFetchIntegrationTests` 和 `DocumentSessionTests` 定向组 57 项通过，0 失败 / 跳过；证据 `artifacts/fetch-validation/network-consent-final.trx`。使用已发布 Native AOT worker 覆盖，测试副本与发布副本 SHA-256 一致。新增覆盖重复加载、同清单源码修改、规范化 / 重排 / 去重、清单缩小或移除、拒绝 / 取消、进行中的旧授权弹窗失效，以及初始化失败后显式重试；保持旧上下文取消、初始化事务回滚和明确停用撤销。

`scripts/test-avalonia-web.ps1 -SkipBuild -NetworkOnly` 23 项专项检查通过，证据 `artifacts/avalonia/web-tests/20261120-095317-eb28d665/result.json`。实际 Avalonia / Monaco 验证加载和编辑后重新加载的弹窗次数、清单变更和停用后的重新授权，以及网络面板和设置开关；授权弹窗截图已检查。新增 `-NetworkOnly` 可独立运行这组检查，避免重复执行不相关的编辑器检查。较宽的 `-ScriptsOnly` 尝试曾在 TypeScript 应用断言处间歇性失败，本次不声称该组整体通过；失败证据保留在 `artifacts/avalonia/web-tests`。未运行完整回归。

`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 已发布并激活 6.1.1，生产自检 36 项通过。`artifacts/avalonia/active-windows.json` 指向 `artifacts/avalonia/releases/20261120-095327-8221d5e5/application`，发布清单 SHA-256 校验匹配，`Start-QuickExcel.cmd` 使用此指针。新增三项真实 worker 回归又针对该包的 Native AOT worker 全部通过，证据 `artifacts/fetch-validation/network-consent-release-aot.trx`；发布、`evidence/worker-native` 与测试副本 worker EXE 的 SHA-256 均为 `EDD21F0011F8B86032A262CC5463170B748C352BFEDF06453C4828240810E80A`。发布仍报告诊断标志未赋值及 Markdown / ODBC / DataSet 裁剪告警，不声称零告警。

### 6.1.0 设置和大正文验证

最终相关定向检查使用生产 Native AOT worker 覆盖：203 项通过，1 项已有 `JSC-MEM-001` 延期检查跳过，0 失败；证据 `artifacts/fetch-validation/network-settings-final-aot.trx`。覆盖 Fetch / Network、Actions、协议和源生成 JSON、JSC 桥、设置迁移，包括实际隔离 worker 在既有 256 MiB 进程限制下的 8 MiB / 10 MiB 上传与下载、超时及禁网后的整轮回滚、精确 16 MiB 帧边界。独立 scripting smoke 通过，未运行完整回归。

设置 / broker / 会话定向组 42 项通过，0 失败 / 跳过，证据 `artifacts/fetch-validation/network-settings-core-final.trx`。覆盖正文默认值与可配置边界、关闭网络开关后的直接拒绝和取消、授权失效、面板正文及元数据预算、动态缩减限额、设置持久化及旧设置迁移。旧文件未包含网络配置时补充默认值，同时保留原有缩放和语言；显式 null 或非法限额仍拒绝。迁移修复的独立复验 4 项通过，证据 `artifacts/fetch-validation/network-settings-migration.trx`。

实际 Avalonia / Monaco 脚本界面 smoke 195 项通过，证据 `artifacts/avalonia/network-ui/20261120-090259-17c2bc96/result.json`。其中设置检查通过真实控件验证 8 MiB / 10 秒 / 200 条 / 50 MiB 默认值、即时关闭网络、授权申请不弹窗即拒绝、应用数值时保持关闭开关、重新读取设置、即时回收旧响应和动态状态栏。已检查同目录 `script-network-settings.png`、`script-network.png` 与 `script-network-consent.png` 的渲染。

脚本声明由 `tools/QuickExcel.ApiGen/src/quickexcel-extensions.d.ts` 重新生成；产品声明 / 能力清单 `--check` 与 ApiGen fixture 自测均通过，保留固定来源校验与 CC BY 4.0 署名。生成物未手工修改。

`ScriptFetchRuntimeTests` 定向组 18 项通过，覆盖 8 MiB UTF-8 请求 / 响应、会导致普通 base64 JSON 膨胀的 8 MiB 二进制请求 / 响应、逐 Action 正文设置、非法限额拒绝、脚本全局变量不能扩大限额，以及原有异步和响应淘汰行为。命令使用 `--artifacts-path artifacts/build-network-runtime`；原生 1 秒执行 / 续接上限保持启用。

发布前 AOT 验证使用 `QuickExcelPublishedWorkerDirectory=artifacts/network-aot-fix/1795166705439-direct-frame/worker`；该发布产物与当次测试目录 worker EXE 的 SHA-256 均为 `8F2B080059A3E11C441D5DED8AECC50EF983845927093ADCE6A932AB06AB09E1`。直接网络帧写入避免了大正文请求的额外包装 JSON 树，AOT 和普通开发 worker 均通过上限检查；没有放宽进程、原生时间或 IPC 预算。此前未通过大正文 AOT 复验的候选包保留在 artifacts，当时已回退启动指针，未作为最终交付使用。

最终交付再次运行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，生产 self-check 36 项通过。包路径为 `artifacts/avalonia/releases/20261120-092829-c6b8709e/application`，版本 6.1.0；`artifacts/avalonia/active-windows.json` 指向该包，清单 SHA-256 匹配，Web 资源为 production 模式，`Start-QuickExcel.cmd` 使用这一已验证指针。

针对最终发布的 Native AOT worker，Fetch 集成及 Actions 81 项复验全部通过，证据 `artifacts/fetch-validation/network-settings-release-aot.trx`，再次覆盖 8 MiB / 10 MiB 上传和下载。最终包、其 `evidence/worker-native` 及复验目录的 worker EXE SHA-256 均为 `06E587653BBB21FE048496B44465A382DD0B7FAB2783570ABC2330021B1A91BF`。发布仍有既有 Markdown / ODBC / DataSet 裁剪告警，生产自检通过；没有将其描述为零告警发布。

### 6.0.0 历史验证与发布

此前 Windows x64 验证：

- 定向测试：247 通过，1 项 `JSC-MEM-001` 已有延期测试跳过，0 失败。证据 `artifacts/fetch-validation/fetch-targeted-final.trx`；包括 Fetch / Network、入口约定、Actions、JSC 桥、控制台、Events / 虚拟文件及模块边界，没有运行完整回归。
- 使用生产 `QuickExcelPublishedWorkerDirectory` 覆盖的 Fetch 集成、入口及 Actions 测试：92 通过，0 跳过 / 失败。证据 `artifacts/fetch-validation/fetch-native-aot.trx`；测试与发布 worker EXE SHA-256 均为 `18DFE8EB40087628F10360F341A23231624785013E176FF305FEE8D810DB87C2`。
- 实际 Avalonia / Monaco 脚本界面 smoke：188 项通过；授权弹窗及网络面板截图已检查。证据 `artifacts/avalonia/network-ui/20261120-084441-0491eaaf/result.json`，同目录 `script-network.png`、`script-network-consent.png`。
- 原生桥完成源码校验和重建；独立 scripting smoke 通过，保留 `JSC-MEM-001` OPEN。API 生成器自测和产品声明 `--check` 通过。
- `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 完成，生产 self-check 36 项通过。新包 `artifacts/avalonia/releases/20261120-084617-84c72312/application` 为 6.0.0；`artifacts/avalonia/active-windows.json` 已指向该包，清单 SHA-256 校验一致。发布包 worker 与上述 AOT 测试的 EXE 哈希一致。旧活动包保留。

以上不是 macOS / Linux 的本次物理验收证据。
