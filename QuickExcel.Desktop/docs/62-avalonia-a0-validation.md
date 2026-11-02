# Avalonia A0 技术验证与基线

2026-11-02。对应 [61 号方案](61-avalonia-migration-plan.md) 的 AV-001–004。
**2026-11-03：A0 工程验证完成，实机验收延期，开始 A1；正式入口仍为 WPF。** 本记录区分已经执行的检查、后续验收协议和仍缺少的证据。

## 1. 实现范围与版本

新增独立 [AvaloniaProbe](../../tools/QuickExcel.AvaloniaProbe/README.md)，不加入产品解决方案或发布脚本。
它在同一进程绘制 Avalonia Fluent 控件、共享单元格、RichTextKit 富文本、五种图片和三类图表，
输出直接 PDF，并嵌入运行本地 Monaco/PDF.js 的 NativeWebView。此探针没有工作簿编辑、文件打开或脚本执行能力。

| 组件 | 本次固定值 |
| --- | --- |
| Avalonia 框架 / Desktop / Fluent | 12.1.4；MIT |
| Avalonia.Controls.WebView | 12.1.1；MIT |
| 共享 SkiaSharp / HarfBuzzSharp | 4.154.0 / 14.2.1.400；保持既有版本 |
| PDF.js / Monaco | 6.5.438 / 0.57.0；Apache-2.0 / MIT |
| esbuild | 0.28.2；MIT，仅构建使用 |
| WebView2 SDK COM 投影 | 1.0.4022.49；包内 Microsoft BSD 风格许可，仅引用 Core，不加载 WPF/WinForms 控件 |
| 本地 SDK / runtime / Node | .NET SDK 10.0.401 / runtime 10.0.12 / Node 22.18.0 |

源码 commit、直接/传递许可、原生依赖边界见 [SOURCES](../../tools/QuickExcel.AvaloniaProbe/SOURCES.md)。
NuGet/npm 锁文件入 Git；实际 30 个 NuGet 包的签名、规范化内容哈希和压缩包哈希已核验，44 个 npm 锁文件条目完整登记。
后者包含未安装的平台可选包，不能把清单数量等同于实际运行时加载数量。生成的依赖表位于每次 Windows 运行的 `dependencies/`。
Skia/HarfBuzz 的 WebAssembly 旧版本传递资产不进入受测桌面进程；实际加载路径、程序集版本和原生文件 SHA-256 保存在 `result.json`。

## 2. 资源、worker 与生命周期决策

- `QuickExcel.Web` 使用 ESM，构建输出位于 `artifacts/avalonia/web`；每次从空的独立暂存目录生成，再替换当前目录，保留旧目录作为证据。不会把旧 bundle 混入新资源清单。
- 213 项资源包含页面、CSS、真实 TypeScript/编辑器/PDF worker、CMap、字体和编解码资产及许可。产品运行不需要 Node；不使用 CDN。
- A0 的统一试验源为 `127.0.0.1` 随机端口、每实例 192-bit 随机路径令牌。服务器只从内存提供已校验 size/SHA-256 的明确资源清单；不映射文件目录。并发上限 16、请求头 8 KiB、请求超时 10 秒。
- Loopback 仅允许精确 Host、HTTP GET 和精确路径；无令牌、穿越、额外查询和错误 Host 返回 404。CSP 限制脚本/worker/connect/font 为同源，禁用远程导航、新窗口、表单和嵌入；Monaco 样式允许内联，PDF WASM 允许编译。
- 探针协议只接受 protocol/instance/type/requestId 明确匹配的 ready/ping/pong/dispose 消息；不暴露 .NET 对象或通用执行入口。A4 仍需独立实现正式文档/模块/修订消息协议及故障注入。
- 关闭先等待页面释放 Monaco model/editor、PDF loading task 和 workers 的 ack，再卸载 WebView、等待 AdapterDestroyed 与 GTK 释放队列，最后停监听并清空资源。运行脚本同时检查进程退出码与 JSON，禁止只看 `passed: true`。

三项真实兼容问题已查明：Windows NativeWebView 需要受支持系统的应用 manifest；WebKit 的 `InvokeScript` 不能把 JS Promise 当作可序列化返回值，单向调用末尾显式返回 `true`；Linux WebKitGTK 必须使用公开的 `X11PlatformOptions.UseGLibMainLoop = true`。
最后一项曾导致 managed 检查全通过、主函数返回后 `SIGABRT`，GDB 栈位于 WebKit/GObject 释放；单独渲染不复现。
启用 GLib 主循环后连续三次完整运行退出 0。没有压掉信号、禁用 sandbox、修改 native 库或回退渲染依赖。
失败结果、GDB 和修复前后退出日志保留在 `artifacts/avalonia/linux-vm/`。

Windows 虚拟 HTTPS 已另行实测：`-VirtualHttps` 通过公开平台 handle 和 WebView2 `ICoreWebView2_3` 映射
`https://a0-<random>.invalid/<random>/`，真实 worker、PDF 提取和双向消息通过。仅映射新建的已校验测试资产目录，跨源访问设为 Deny，关闭清除映射；不映射工作簿目录。
Microsoft 的[虚拟主机 API](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_3)支持 worker；
[资源拦截文档](https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/webresourcerequested)同时说明该映射不触发 WebResourceRequested。
因此该实验使用 CSP meta，而非自定义响应头；A0 保留 loopback 为两平台统一探针，A4 的 Windows 适配可以采用虚拟 HTTPS，但仍需实现动态 PDF、过期身份、响应头/隔离及清理的完整契约。
两种模式的结果与限制分别记录，不能把 loopback 的拒绝测试外推为虚拟映射的拒绝测试。

## 3. 已执行证据与平台边界

| 检查 | 结果与证据 |
| --- | --- |
| 现有 Release 解决方案 | 0 警告、0 错误；未改变生产项目 |
| 共享 Rendering.Test | 20/20；`artifacts/avalonia/tests/a0-rendering.trx`，没有运行全量 Core 回归 |
| Windows 自动探针 | `artifacts/avalonia/20261102-144952-ece89fe2`；Avalonia/共享 Skia/PDF、真实 workers、双向消息及资源拒绝检查通过，进程退出 0 |
| Windows 桌面补测 | `artifacts/avalonia/20261102-145938-346778b3`；实际系统拼音候选、空格提交“你”、原生输入框/Monaco 焦点往返、分隔条拖动、裁剪及跨 WebView 的原生菜单覆盖；关闭退出 0。截图和提交后的 `desktop.json`、`web-disposed.json` 保留 |
| Windows 虚拟 HTTPS | `artifacts/avalonia/20261102-152514-18b95340`；两个 worker URL 均为本地虚拟 HTTPS，检查通过、退出 0；构建 0 警告/0 错误，30 包签名/内容核验通过，无 WPF/WinForms SDK 载荷 |
| Linux 自动探针 | Ubuntu 24.04.5 x64、WebKitGTK 2.52.6、普通 probe 用户、Xvfb；修复后 results-13/14/15 三次退出 0，`repeat-04.log`；非真实桌面 IME 证据 |
| 跨系统 PDF | Windows 上述自动运行与 `linux-vm/results-15` 的三页报表经 MuPDF 1.26.7 独立检查通过；分页和提取文本一致，指定文字几何差异 < 0.5 pt，字体嵌入、矢量内容、页面归属及逐块视觉比较通过 |
| 最终源码复验 | Windows loopback `20261102-154303-0d47fa45`、HTTPS `20261102-154310-cd811d33` 均退出 0；WebView2 156.0.4314.58；Linux 最终载荷连续三次退出 0，`linux-vm/final-linux.log`，最后一次完整结果在 `linux-vm/final-results`，含每个 native 文件 SHA-256 和独立跨系统 PDF 验证 |
| 验证器/脚本 | PowerShell 解析、Node 语法、Linux bash 语法通过；`--report-only` 在两平台输出通过，原默认 text+report 路径用复制到 `artifacts/avalonia/validator-legacy` 的既有样本复验通过，没有覆盖原始证据 |

Windows 探针实际 RenderScaling 为 **1.25**，已记录在 `desktop.json`。本次只声明测试过的系统与内核组合，不将 Ubuntu 24.04 默认最初发行的 WebKit 版本、所有 Linux 桌面、Wayland 或 macOS 视为通过。
虚拟机来自既有测试镜像的新 qcow2 overlay；只修改仓库 artifacts 中的独立副本。WebKit 以普通用户运行，未禁用系统沙箱。
Xvfb 使用软件图形，不证明目标 GPU 性能。真实跨屏 DPI、辅助技术和物理打印仍是 A3/A5/A6 的独立验收。

Linux 图形会话已安装 Xorg、Openbox、fcitx5/Pinyin、Noto CJK，在 1440×900 显示上实际打开内嵌 WebKit，鼠标/焦点/页面绘制可观察。
然而通过当前 Windows→QEMU 自动输入通道发送的字母、Enter 和快捷键未能改变文本；GTK/SDL 两种虚拟机显示后端、USB 键盘和独立 GTK Entry 对照均未取得有效输入证据。
屏幕键盘对照也没有得到可用提交结果。2026-11-03 用户实测反馈：“无法输入中文，没有出现中文输入法候选框”。**Linux 中文输入记为已知未解决问题，不能记作通过，也不能仅据此归因为 Avalonia 缺陷。** AV-003 的真实桌面输入验收延期。
本地 `linux-vm/` 保留 VM 启动/安装日志；会话内 `desktop-03`、`fcitx.log` 与独立 GTK 对照用于继续定位。GTK Entry 和屏幕键盘均为测试工具，不进入仓库产品源码。

## 4. WPF 功能与性能基线

源码基线为 `37911cb04374b3dabd4413f8b7993c1c523c2eb0`（方案提交），产品实现为其父 `4a0af74`。
保留的受测 WPF 包为 `artifacts/publish/20261102-130713-e1444da9`；该包原始 metadata 如实记录构建时 `35877e7` 和 dirty 状态，不能事后改写成干净提交构建。
本次冻结全部文件 SHA-256 与原 metadata 至 `artifacts/avalonia/baseline-20261102/wpf-package-identity.json`，
清单 SHA-256 为 `A0AB6B477677EF02B6A1A084DC9B4A59688E5C738DE14527ED3F7793639AB285`。

| 对齐组 | 现有入口及迁移后需要重放的证据 |
| --- | --- |
| 文件、关闭、保存冲突、恢复 | MainWindow.Workspace/Shell；Workspace 模式、SaveChecked/指纹/关闭取消与独立设置目录 |
| 网格、视口、选区、编辑所有权 | Interaction/GridResize/Zoom/Views/FormulaEditing；Interaction、GridIntegration、GridScrollRange、OversizedGridCell |
| 剪贴板、填充、结构、附件 | Clipboard/Fill/SheetManagement；40–43 号契约、一次撤销、隐藏/合并/保护及富文本/链接/批注拒绝规则 |
| Ribbon、格式、统计、自动尺寸 | Ribbon/FormatEfficiency/SelectionStatistics/CellText/ColorPalette；Ribbon、FormatEfficiency 模式及 49 号数值格式基线 |
| 筛选、排序、表、验证、条件格式、清洗、分组 | DataOperations/Tables/Validation/ConditionalFormatting/DataCleaning/Outline；50–56 号 Excel/存储/桌面对应矩阵 |
| 查找、工作表与名称 | Search/SheetManagement；42/43 号导航、名称和引用调整基线 |
| 图片、图表、分页、预览、打印 | ReportObjects/Reporting；Reporting 模式、E4 五页 fixture、三类图表、图片锚点、保护与页设置 |
| 脚本、模块、授权、签名、日志 | Scripts/Signatures；v1/v2/IPC4、immutable snapshot、授权撤销、worker 隔离及原子 actions；A4 增补 flush/ack 竞态 |

该表是迁移重放清单，不声称新增宿主已经实现这些功能。60 号记录与各专题的已有结果保留，不把本次未重跑的历史测试写成本次通过。

本次重新顺序运行同一 WPF 包，基准期间测试 VM 关闭，窗口 1400×850 DIP，每场景预热 5 次、采样 60 次：

| 基准 | 本次结果 | 证据 |
| --- | --- | --- |
| 横向输入到 WPF CPU 绘制完成 | 普通格 p95 **14.9341 ms**，样式格 **12.4028 ms**；≤16.67 ms，全场景通过 | `artifacts/validation/m5-grid/20261102-145012-97c5f4db`，包含全部 raw phases |
| 拖选与自动滚动 | 自动滚动 120 次更新、间隔 p95 **19.0365 ms**；≤33.34 ms，全场景通过；纯选区模型读取 0 | `artifacts/validation/m5-grid/20261102-145055-947ac45c` |

继续使用 25k/100k/250k 存储格、8,192 样式、稀疏远端、冻结/隐藏/合并/超大格矩阵；百万上限不冒充稠密百万格性能。
Avalonia 后续必须实现相应的内容/覆盖层绘制完成观测点，报告输入、状态应用、内容绘制与提交时刻；这些 CPU 指标均不代表物理呈现延迟。

## 5. 新 Web 负载与初始验收预算

以下在正式 A4 接入前冻结，是**待执行的性能验收协议**。A0 的一页 fixture 和一次页面内部 elapsed 值不代表已满足此表。
使用独立 profile、相同包/字体/文档/窗口和 100% 缩放，记录 CPU/GPU/系统内核；冷启动 5 次、预热 5 次后交互 60 次，保留全部样本，不与其他测试并跑。

| 负载 | 测量端点及初始阈值 |
| --- | --- |
| WebView 冷启动 | 从宿主创建请求到 ready/worker 就绪 p95 ≤ 5 s；browser/runtime 缺失明确失败，不能计作快速成功 |
| E4 五页、100 页、2,000 页轻内容 PDF | 已生成文件交付到首张可见画布完成 p95 ≤ 5 s；不把 PDF 生成时间藏入这个指标，另报生成/取消延迟 |
| 高图像密度单页 PDF | 4 张 2048×2048 RGBA 图像、共 64 MiB 解码输入；首张画布 p95 ≤ 5 s，缩放后可见画布完成 p95 ≤ 250 ms |
| Monaco 64 模块、总源码 1 MiB UTF-8 边界 | 普通输入到确认绘制 p95 ≤ 50 ms，语言服务补全 p95 ≤ 500 ms；超限仍按既有源码预算拒绝 |
| 进程树资源 | 桌面+浏览器全部子进程峰值私有提交：五页/普通编辑 ≤ 768 MiB；2,000 页/高图像密度 ≤ 1.5 GiB；不能只采 .NET 父进程 |
| 回收与并发交互 | 打开/关闭 10 轮、空闲 30 s 后相对预热基线增长 ≤ 64 MiB，无遗留页面/worker；PDF 滚动/缩放、Monaco 补全期间主网格继续满足既有 16.67/33.34 ms 预算 |

100/2,000 页使用固定页面序号和轻内容，每页相同 A4 尺寸、少量矢量格和中文/拉丁记录；页数/内容/hash 随样本生成器登记。
A4 实施时需补足固定 fixture 生成与进程树测量工具。若预算失败先定位，再按实际证据重新评审，禁止在结果出来后悄悄抬高阈值。
真实设备性能与 VM 的可行性结果分别记录。

## 6. 复现与后续门禁

Windows 构建/运行、Linux 独立载荷、退出码检查和跨系统独立 PDF 命令见 [探针 README](../../tools/QuickExcel.AvaloniaProbe/README.md)。
所有 managed 输出使用仓库内 `--artifacts-path`；Web、字体、浏览器 profile、VM、日志和验证工具均留在 ignored artifacts。

2026-11-03 用户明确调整实施顺序：“涉及实机验证的先跳过，在实施日志中记录即可，后期统一验证。”据此完成 A0 工程检查并进入 A1；无需等候实机条件再继续代码实施。该调整延期的是实机验收，不表示缺失证据已经通过，不降低功能、数据或隔离要求。

统一延期清单：Linux 中文输入候选/提交/取消与 WebKit 焦点/弹窗，macOS 桌面和 WKWebView，跨屏 DPI、真实辅助技术、物理打印、干净目标设备的部署和性能。后续阶段继续记录各自的自动检查、已知问题和待验收项；用户验收前不推送远程。
非 Windows worker 沙箱、原生桥、Keychain/密钥服务、系统打印、签名/公证与干净环境仍属 A6；当前探针不执行脚本，不能视为这些能力的证明。
M4 严格策略/M6 OS 安全、完整部署与恢复、M5/M6 实机输入/辅助技术、E5 继续按原专题保持未关闭。
