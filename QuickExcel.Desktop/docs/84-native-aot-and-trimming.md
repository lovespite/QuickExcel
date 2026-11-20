# 4.0 Windows Native AOT 与发布裁剪

## 范围与产物

产品主版本由 `QuickExcel.Desktop.Avalonia.csproj` 提升为 `4.0.0`，ClickOnce
使用同一版本的四段形式 `4.0.0.0`。Windows x64 Worker 单独 Native AOT
发布，桌面仍为自包含 .NET 10，以 `TrimMode=partial` 裁剪支持裁剪的程序集。
未声明裁剪兼容的应用和第三方程序集保留；不通过删除未知 `System.*` DLL
来猜测依赖。Linux/macOS 保留原有经过独立验证的托管 Worker 发布路径，
本次 Windows 验证不代表其他平台的 AOT 验收。

`scripts/build-avalonia-release.ps1` 在独立证据目录发布 Worker，再通过
`QuickExcelPublishedWorkerDirectory` 将其载荷交给桌面构建。生产 Windows
发布缺少该参数时失败，避免静默混入普通 apphost/完整 CLR。测试也可用同一
参数运行真正的 AOT Worker，而普通无 RID 的开发构建仍便于调试。
桌面与 Worker 的构建缓存和 NuGet 锁目录独立；还原显式传入
`RuntimeIdentifier=win-x64`，使 AOT/裁剪条件与 publish 一致。两份实际锁
都进入包的 metadata，避免 `restore -r` 只扩充 RID 图而漏还原条件工具包。

主程序目录不再复制 `JavaScriptCore.dll`、ICU、VC++ 支持库、
`typescript.bundle.json` 和 JSC 身份文件。Worker 仍拥有全部依赖；
AppContainer、只读私有载荷、Job、句柄白名单、持久上下文、崩溃恢复和
执行/编译资源预算继续使用既有实现。JSC 的动态脚本执行与 .NET AOT 相互独立。

`test-windows-payload.ps1` 同时供正式包和 ClickOnce 验证调用，检查原生 x64
Worker、没有 CLR/托管入口残留、只有一份引擎资源及生产 JSON 反射已关闭。
生产自检继续执行脚本、TypeScript 编译、事务、持久上下文、Web、打印和恢复。
JSC 源码包、许可证、修改记录和重建说明全部保留。

## JSON 兼容性

IPC 固定 DTO、编译结果、脚本桥固定消息、设置、信任/签名偏好、编译缓存、
恢复记录及 Web 信封使用 System.Text.Json 源生成元数据。动态的脚本值和
模型/Web 对象使用明确的 JSON 值/字典写入器，不对匿名对象做反射序列化。
生产 Worker 与 Windows 生产桌面关闭默认 JSON 反射。

`JsonArray.Add<T>` 也会隐式请求反射元数据，不能因参数是字符串就当作安全。
存储清单采用 `Add(JsonNode)` 与明确的 `JsonValue.Create(string)`；生产候选包
自检捕获了 localStorage 保存分支的这一问题，修复后另用反射关闭的测试进程
覆盖本地存储、虚拟文件清单、Apply、保存重开及损坏包拒绝。

保留 IPC 11、文件版本、原有字段顺序/大小写、转义、null 与数字语义；摘要和
签名输入不能因改用 DTO 而改变。源生成回归用旧版序列化作为独立字节对照，
覆盖协议/摘要及持久化设置。重复键、未知字段、必需字段、非空注解和现有
16 MiB 消息限制仍有效；失败不能发布部分事务。

## 工具链与验证

Windows 发布要求 .NET SDK `10.0.404` 或更新版本；发布脚本优先使用
`artifacts/tools/dotnet-10.0.404` 下解压的官方 SDK，也可使用机器安装的新版。
本次本地 SDK ZIP 按微软 release-metadata 提供的 SHA-512 校验后解压，
不改动机器已有 SDK。Windows 自包含运行时和 Native AOT 编译器固定为 .NET `10.0.15`，由
`Directory.Build.props` 的 `QuickExcelWindowsRuntimeVersion` 统一配置，
实际 NuGet 锁随发布证据保存。SDK 默认 `10.0.12` 的编译器在试发布时报告
NU1903；本次选择修复后的补丁版本，未屏蔽漏洞审计。
微软公告：<https://github.com/dotnet/runtime/security/advisories/GHSA-p3jp-9v83-76w9>。

主程序部分裁剪保留并显示 Markdown.Avalonia.Tight 12.0.2 的 8 条上游分析
警告，没有全局抑制。已按包的源提交
`f3f65f3f82b44f569958f0116c188e7a8e60686f` 核对：2 条用于调用栈定位相对图片
（其中旧 BitmapLoader 未使用），3 条位于 Tight 内没有调用者的跨程序集反射
工具，1 条用于显式插件的依赖实例化（本应用未注册插件），2 条用于未随包
分发的可选 SyntaxHigh 插件探测（缺少时捕获异常）。该库未声明 IsTrimmable，
partial 模式保留程序集和资源；当前许可证页直接引用标准解析器与 GithubLike
样式。生产自检实际打开 About 页检查 Markdown 内容与布局。以后引入图片、
第三方插件或改用 full 裁剪时，必须重新审查，不能沿用这次结论。

发布脚本在 finally 中恢复调用者的 PATH、DOTNET_ROOT 和 OS，避免本地
仅含 .NET 10 的 SDK 环境影响后续依赖 .NET 8 的 Mage。通过包装脚本分别在
原始环境、未设置 DOTNET_ROOT/OS 的环境触发已有输出目录错误，验证环境
完整恢复、活动指针未变、没有新建发布目录，且 Mage 可正常启动。

### 2026-11-20 本地交付验证

只运行受影响范围的测试，没有运行全量回归：

- IPC/源生成/消息预算 51 项通过；桥接、模块、TypeScript 等 89 项通过，
  1 项既有 JSC-MEM-001 延期测试跳过；关闭 JSON 反射的 10,000 次调用冒烟通过。
- 宿主 JSON、Web、设置/信任/恢复相关定向测试通过；额外存储兼容测试
  106 项通过；关闭 JSON 反射的存储/虚拟文件测试 5 项通过；打印 8 项通过。
- 最终已签名包中的 AOT Worker 用于 386 项真实进程测试，全部通过、无跳过。
  范围包含协议、沙箱、故障、恢复、动作、批处理、编译器、事件、存储、
  API P1–P5、模型与签名。测试输出的 Worker SHA-256 与活动包一致：
  `8A980957C0BF986549C3124551278970C22025389F8B8167D25EBC5B454B7FC0`。
  TRX：`artifacts/aot-v4-evidence/final-native-worker.trx`。
- RibbonFilter 843 项 UI 检查、Drawing 37 项检查通过，证据分别为
  `artifacts/avalonia/workspace/20261120-023304-8b8efdb2` 和
  `artifacts/avalonia/workspace/20261120-023332-262fa9fd`。
- 正式包、`Start-QuickExcel.cmd`、ClickOnce 实际载荷均通过 34 项生产自检，
  包括新增的设置页 Markdown 渲染检查。启动器证据：
  `artifacts/avalonia/launcher-self-check-20261120-0252`。

已运行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，使用既有代码签名
证书签名。`artifacts/avalonia/active-windows.json` 指向
`artifacts/avalonia/releases/20261120-024227-929d0d3c/application`，
release.json 哈希已核对：
`81579BF63EE1ED7BF0E8382E887BFD43B9FD9785C60A2C478616753C150EFB41`。
Worker 原生发布没有 AOT/裁剪警告；主程序自身裁剪警告已解决，保留上述
8 条经审查的上游警告。运行环境为 Windows 10.0.22621、WebView2
157.0.4341.60、.NET 10.0.15。

与保留的 `20261120-020817-5493b6da` / 3.13.0 包比较（MiB = 1,048,576 字节）：

| 项目 | 3.13.0 | 4.0.0 |
| --- | ---: | ---: |
| 完整发布目录 | 454,813,842 字节 / 433.74 MiB | 245,606,935 字节 / 234.23 MiB |
| 文件数 | 754 | 431 |
| ScriptWorker 目录（含原生引擎） | 154,353,868 字节 / 147.20 MiB | 78,339,521 字节 / 74.71 MiB |
| 主程序 System.Private.CoreLib.dll | 16,033,576 字节 | 3,265,024 字节 |
| ClickOnce ZIP | 232,914,231 字节 / 222.12 MiB | 144,702,889 字节 / 138.00 MiB |

发布目录减少约 46.0%，ClickOnce ZIP 减少约 37.9%。上述大小保留了全部
许可证与 JSC 源码归档，不以删除开源合规材料换取体积。

ClickOnce 为 `4.0.0.0`，沿用原证书；433 项应用清单文件哈希、两个清单
XML 签名、代码签名、前置组件及生产自检均通过。没有改动证书信任存储。
证据：`artifacts/clickonce/work/20261120-024702-f2315b6d/verification`。
完整分发包：
`artifacts/clickonce/QuickExcel-4.0.0-win-x64-20261120-024702-f2315b6d.zip`。
归档的 440 个文件逐一核对大小和 SHA-256，并检查分发清单全部 439 个记录；
结果在 `artifacts/aot-v4-evidence/clickonce-zip-verification.json`。
ZIP SHA-256：
`398854FF7CDC85972C91F0D6B5EC242293E40F09A0A58B488600A3CF91868C5D`。
旧活动包与旧 ClickOnce ZIP 均保留。

物理 IME/DPI/辅助功能/打印、干净设备 ClickOnce 安装、平台隔离最终验收及
JSC-MEM-001 延期状态不受本次优化影响。
