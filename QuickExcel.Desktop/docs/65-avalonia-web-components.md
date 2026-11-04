# A4：Monaco、PDF.js 与原生打印接入

2026-11-03。**A4 工程实现与专项完成，实机项按用户要求延期。** 本文保留 A4 阶段证据；当时默认入口尚未切换，后续已在 A5 切换为 Avalonia，并在 A7 移除 WPF 生产项目。各阶段提交已在迁移分支推送，当前状态见 [68 号收尾记录](68-avalonia-retirement.md)。

用户已要求实机验证先跳过、记入日志后统一验收。Linux 中文输入无法出现候选框的反馈继续记为未通过；物理打印、跨屏 DPI、真实辅助技术、Linux/macOS 桌面等保持延期，不以自动检查替代。

## 当前实现

- 2026-11-03 脚本编辑辅助更新：补齐真实补全框、悬浮类型/说明、函数参数提示和当前模块错误导航；加入 `$` API 别名。使用方法、专项结果与范围见 [71 §12](71-jsc-typescript-implementation-log.md#12-脚本编辑辅助与-api-别名)。
- 正式 Web 资源使用锁定的 Monaco 0.57.0、PDF.js 6.5.438、esbuild 0.28.2。清单逐项验证大小与 SHA-256；独立随机 loopback origin 仅服务白名单资源和有租约的 PDF，不暴露工作簿目录。普通消息限制 256 KiB，源码总量仍为 1 MiB，PDF 不经过 JSON/Base64 桥。
- 主程序为每个页面维护文档身份、视图代次、请求标识和消息能力集；拦截外部导航与弹窗，使用心跳发现失败。关闭先请求页面释放，再拆除 NativeWebView，等待在途资源读取，最后删除自有临时 PDF。Linux GLib 销毁路径沿用 A0 验证方式，正式页面的 Linux 两轮回归已通过（见最终证据）。
- 脚本工作区使用原生模块、角色、授权、操作、日志和签名控件，内嵌 Monaco。稳定 model URI 保留普通模块切换中的撤销栈、光标与滚动位置。添加模块更改模型集合时重新创建页面代次。静态服务只配置 ECMAScript 与 QuickExcel 声明、已声明的逻辑模块，无 DOM/Node 库；不执行用户源码。
- 编辑消息立即撤销旧授权；按版本、顺序与 SHA-256 确认分块。保存、运行、角色、签名与关闭先 flush；关键操作期间锁住编辑器，提交一次共享命令。无变化不创建文档命令。页面故障阻止保存旧源码，用户可显式恢复主程序已确认草稿，或放弃未提交草稿关闭失败页面。
- A4 复用原有 Windows AppContainer worker、当时的 IPC 4、动作事务和 DPAPI 证书信任。一次允许不保存信任；始终信任只写独立 Avalonia 用户目录；角色/源码变更撤销授权与签名。后续 A6 已提供受测 Linux 隔离，macOS 执行仍关闭；当前事件/虚拟文件功能使用 IPC 5，见 [69 号契约](69-script-events-and-virtual-files.md)。
- `PrintSession` 拥有冻结 PrintJob、生成任务、取消和临时文件租约；完整 PDF 关闭并经过独立解析验证后才交给 PDF.js。导出重用同一任务，保留同目录临时文件、flush、验证及原子替换。文档改变后旧预览明确标记，用户显式重建。
- PDF.js 使用现有 Viewer/PageView、文本层和 FindController。提供连续页、缩略图、页码、查找、选择复制、整页/页宽与缩放；禁用表单、批注编辑、脚本和外链。自定义有界渲染队列只控制调度及销毁，不修改 PDF.js 私有缓存字段；最多 8 个完整页画布、64 MiB，缩略图另限 12 个 112×160 画布。
- Windows 原生 `PrintDlgExW` + GDI 输出沿用同一 PrintJob，最高 300 DPI，大纸张按 64 MiB 下调栅格 DPI。页范围/份数交给系统对话框；页尺寸与方向必须匹配预览，硬边距不足明确拒绝；通过 AbortProc 和页间检查取消。尚未发送物理打印任务。

## 已运行的检查

| 检查 | 本次结果与证据 |
| --- | --- |
| 草稿协议 | 10 项通过；`artifacts/avalonia/tests/a4-web-drafts.trx`，覆盖旧身份/代次、重放/乱序、哈希、总量、同步竞态、故障恢复 |
| 本地资源与 PDF 租约 | 4 项通过；`artifacts/avalonia/tests/a4-web-origin.trx`，覆盖篡改、路径/来源拒绝、Range、撤销、在途关闭 |
| 打印会话 | 3 项通过；`artifacts/avalonia/tests/a4-print-session.trx`，覆盖冻结导出、取消保留旧文件、读取租约与临时文件清理 |
| 打印坐标/预算 | 2 项通过；`artifacts/avalonia/tests/a4-printer-layout.trx`，覆盖物理原点、尺寸拒绝与大纸张 DPI 上限 |
| 真正 Avalonia + WebView2 页面 | 36 项通过；`artifacts/avalonia/web-tests/20261102-191139-1bf99b01`，包含实际语言 worker、导入/补全/诊断、源码保存/重开、模型撤销、worker actions/签名信任、故障恢复、PDF 文本/搜索/缩放/冻结导出和 64 模块/1 MiB 初始化 |
| 构建 | 当前 Avalonia Release 定向构建通过，0 警告、0 错误 |

协议后续增加单调时钟传输超时与正式包拒绝测试资源，共 **17 项通过**，见 `artifacts/avalonia/tests/a4-web-contracts-2.trx`。

## 独立性能：已完成的工程样本

编辑器首轮：`artifacts/avalonia/web-benchmarks/20261102-191643-5ea243b6`。64 模块合计恰好 1 MiB；预热 5 次、交互采样 60 次。输入到宿主 SHA-256 确认和 requestAnimationFrame 回调 p95 **13.7 ms**，语言服务补全 **4.3 ms**；五次页面打开（同一宿主）p95 **1442.50 ms**。完整进程树峰值私有提交 **617,078,784 字节**，10 轮开关并空闲 30 秒后增长 **18,276,352 字节**，仅余宿主进程。五个独立冷启动进程的检查仍待补足；页面帧回调不是物理屏幕呈现。

PDF 修正后的首轮：`artifacts/avalonia/web-benchmarks/20261102-192332-8b8ca57c`。使用 E4 原始三个 PDF 场景，共 1+2+2 页；100/2,000 页固定轻内容工作簿；四张 2048² RGBA 图片，共 64 MiB 解码输入。每场景五次交付，图片缩放另预热 5 次、采样 60 次；全部原始时间、PDF 哈希与进程树样本保留。

| 场景 | 完成 PDF 交付到首张画布 p95 | 进程树峰值私有提交 |
| --- | --- | --- |
| E4 quickexcel / native / paged | 400.99 / 296.18 / 310.27 ms | 802,738,176 / 759,324,672 / 759,074,816 字节 |
| 100 页 | 306.91 ms | 834,441,216 字节 |
| 2,000 页 | 1078.28 ms | 993,406,976 字节 |
| 高图像密度 | 176.77 ms；可见页缩放 p95 **48.4 ms** | 1,320,398,848 字节 |

达到当前分档预算，但 E4 首场景的 768 MiB 内存余量很小，最终候选必须复核。PDF 生成与解析验证时间在结果中单列，没有混入首张画布时间。正常复用应用浏览器 profile；独立冷启动另测。

初次 PDF 测量 `20261102-192126-3aac45db` 保留为失败：夹具将一张轻内容页误标为 E4 quickexcel，且高图像区域横向产生两页；已改为历史受测 `ui-report.xlsx` 与明确列宽。该轮还在普通场景前预分配高图像夹具、连续启动新的浏览器 profile，导致普通 PDF 进程树峰值超过预算。修正为按场景准备输入、使用实际应用的 profile 复用方式；没有提高阈值或删除失败样本。

真实页面首批 19 项与 26 项结果分别保留在 `20261102-185617-92256d24`、`20261102-190338-35258509`。早期失败样本也保留：JS 赋值隐式创建名字导致测试断言不适当，改为调用缺失的 DOM/Node API；测试消息的 Unicode JSON 转义超过普通消息预算，调整为仍覆盖多分块的夹具。没有提高产品容量限制。

视觉检查发现系统深色偏好使 PDF.js 工具栏文字与白底对比不足，已显式设置浅色配色并加入真实计算样式断言。Monaco 原生工具栏、行号与语法高亮已查看；更完整的辅助技术及物理输入仍延期。

正式资源构建不包含 A0 探针或测试操作。`scripts/build-avalonia-web.ps1 -Probe` 单独构建旧探针；`-Tests` 构建故障注入资源，正式工作区拒绝加载测试清单。`scripts/test-avalonia-web.ps1` 使用独立 artifacts 设置目录。手动视觉检查入口 `--manual-check <artifacts目录> [工作簿]` 同样隔离设置。

## 延期验收与后续阶段

实际 Windows 打印对话框、设备能力及物理打印按用户的实机延期安排留待统一验证。A5 继续候选发布与入口切换；A6 单独实现平台适配，不把 Linux 页面检查扩展为完整平台支持。

## 收尾专项与发现

- 冷启动五个独立宿主、各自全新 profile：`20261102-193410-ac2bc3f3` 至 `20261102-193419-abe69922`，64 模块 Monaco/语言 worker 就绪 p95 **1158.75 ms**；PDF `20261102-193421-ba0b02c5` 至 `20261102-193432-d2906591`，浏览器 ready **383.98 ms**、100 页 PDF 交付 **328.79 ms**。不同冷进程与同一进程反复开关分别记录。
- 100 页 PDF 十轮开关：`20261102-193239-a71a6fa6`，30 秒后仅余宿主，私有提交增长 **6,344,704 字节**。最终采样器修正后的结果另行补记。
- 并发：`20261102-193354-4548f794`，真实 64 模块补全与高图像 PDF 缩放每轮间隔 100 ms，同时原网格 60 样本滚动、拖选、25k/100k/250k 负载。普通/样式绘制 p95 **7.2965/9.2522 ms**，拖选 **18.729 ms**；补全 **1.8 ms**、缩放 **45 ms**。2560×1540 工作区，125% DPI，两个浏览器部分可见且置顶，窗口位置随原始数据记录；不是物理输入/呈现证据。
- 并发初轮 `20261102-193148-30383b65` 暴露未绑定尺寸轴的鼠标悬停除零，已修复；**6 项** ViewportAxis 回归在 `a4-unbound-grid.trx`。下一轮原生悬停重写了合成拖选位置，致零帧；基准仅在合成拖选期间关闭控件的物理命中，继续使用同一生产动画调度，不改变产品拖选逻辑。失败 JSON 对空样本也保留 Infinity，不再丢失证据。
- 最终 Windows 功能检查增加原生导出后备、PDF 页故障/重建、受阻公式拒绝打印、API v1 保留、实际只读数组公式工作簿：**41 项**，`artifacts/avalonia/web-tests/20261102-194459-0ca97e5a`。打印生成中取消新增检查，与租约/导出/坐标共 **6 项**，`a4-print-session-final.trx`。只读夹具初次误将含脚本主文件保存为 XLSX，被既有合同正确拒绝；已修正夹具，产品规则未放宽。

内存复核失败保留于 `20261102-193704-ea70f82d`、`20261102-193833-06e0558f`。第二轮逐进程记录发现采样器只按 PID 跟踪时，旧父进程 ID 被重用，会把更早启动的 VBCSCompiler 误纳入子树。修正为保留进程句柄到退出、校验子进程创建时间，并记录每进程名称/父 ID/私有提交；不按进程名过滤负载。

普通 E4 的中文字体 PDF 达 13–23 MiB；浏览器原生 FontFace 转换与加载造成另一部分峰值。PDF.js 改用公开 `disableFontFace` 的字形路径绘制，仍使用嵌入字体、保留文本层/搜索，不回退系统字体或图片化导出。最新两轮全部性能阈值通过：`20261102-194116-c621cf57`、`20261102-194706-fc2ef1e7`。后一轮 E4 三场景峰值 **803,180,544 / 786,169,856 / 689,213,440 字节**；100 页 **796,651,520**、2,000 页 **963,428,352**、高图像 **1,312,735,232** 字节。首张画布 p95 **388.66 / 290.09 / 294.02 / 265.42 / 393.74 / 125.56 ms**；图片缩放 **49.1 ms**。E4 最坏值仍接近 768 MiB，发布候选需保留这项余量风险。机器字体文件 SHA-256、CPU、GPU 驱动信息写入 `machine.json`，没有提高预算。

实际 PDF.js 五页 canvas 在 `20261102-194436-15673a16`，由 MuPDF 1.26.7 独立比较。校验使用 PageView 实际 viewport/outputScale；最初用四舍五入后的画布宽高反推比例产生假位移，已修正。表格描边与单元格裁剪交界的灰度抗锯齿有差别（原通道法最大 13.74%，原值仍写入结果）；灰度以一像素墨迹几何比较，彩色内容继续用 64 通道邻域容差，128 像素占用块门槛仍为 10%。五页最大几何/颜色错配 **0.3351%**。这是渲染器间抗锯齿差异处理，不代表逐像素一致；截图、画布变换、独立渲染及结果均保留。脚本 `verify-avalonia-pdf.py` 可复现。

原生打印实现依据：[PRINTDLGEXW](https://learn.microsoft.com/en-us/windows/win32/api/commdlg/ns-commdlg-printdlgexw)、[DEVMODEW](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-devmodew)、[StretchDIBits](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-stretchdibits)、[AbortProc](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nc-wingdi-abortproc)。PDF.js 队列基于已锁定包的公开 Viewer/PageView 接口；后续升级需重跑专项。

## A4 最终工程证据

- Windows 正式 WebView2：**42 项通过**，`artifacts/avalonia/web-tests/20261102-200531-3181c3bc`。最终解决方案 19 个项目 Release 构建 **0 警告、0 错误**；协议/资源 17 项、打印会话/坐标 6 项、未绑定轴 6 项的定向回归均通过。
- 最终 Monaco：`20261102-194802-9018add1`，输入确认 p95 **13.2 ms**、补全 **4.9 ms**、五次页面打开 **1034.28 ms**；峰值 **616,140,800 字节**，十轮关闭后增长 **28,794,880 字节**。PDF 十轮关闭：`20261102-194919-0c06e6ec`，增长 **−1,585,152 字节**；均只剩宿主。
- 最终并发：`20261102-195027-0b2ef22a`，网格普通/样式/拖选 p95 **6.5813 / 10.6496 / 19.3901 ms**，真实补全 **2.9 ms**、PDF 缩放 **49.2 ms**，原预算通过。
- 独立视觉负向校验 `artifacts/avalonia/visual-negative-a4` 在派生图片中擦除表格，错配 **79.35%**，按预期失败。原始五页截图未覆盖；原通道抗锯齿差异仍随独立结果保存。
- Linux Ubuntu 24.04.5、WebKitGTK **2.52.6**、普通 `probe` 用户、Xvfb 1440×900：`artifacts/avalonia/a4-linux/results-05` 与 `results-06` 连续两次各 **35 项通过**，退出 0，结束后没有该用户的浏览器/宿主残留。没有关闭 WebKit 沙箱。Windows AppContainer actions 与 DPAPI 明确跳过，非 Windows 脚本仍拒绝执行。入口 `scripts/test-avalonia-web-linux.sh` 记录二进制与 Web 清单哈希，拒绝 root、无图形会话或禁用沙箱。

Linux 首轮 `results-01` 在 `Opened` 回调过早创建 WebView，GLib 主循环尚未启动，造成等待。主窗口启动工作延后至 Dispatcher Background 后解决。`results-02` 至 `results-04` 暴露 Noto CJK 同形字映射：Skia PDF 以 ActualText 保留“文”，未修改 PDF.js 却提取成康熙部首“⽂”；MuPDF 独立提取确认为原文。初次只调整复制检查不足以解决，失败保留。

正式修复是对锁定 PDF.js worker 的文本提取作有限补丁：处理 Skia 单个字形簇的内联 ActualText，保留位置与原始 Unicode，不做全局 NFKC。构建核对上游 SHA-256 和唯一修改位置，随包附 Apache-2.0 许可及修改声明。独立手写 PDF 验证同形异字、连字、多字符、嵌套/空 span 和普通文本；真正文本层复制及中文查找在两平台通过。实现与范围见 `QuickExcel.Web/patches/pdfjs-actualtext.md`。不宣称通用 tagged PDF、跨 Form XObject 或多行替换选择支持；预览仅接受本程序打印会话 PDF。

Linux Xvfb 不是物理桌面、真实辅助技术或中文 IME 证据。此前用户“无法输入中文，没有候选框”仍未关闭。macOS、实际打印、跨屏 DPI 等按用户要求统一后验；既有 M4/M6/E5 未关闭事项保持原状态。

最终 ActualText 补丁版本复测 `20261102-200753-dc7c4912` 的普通 E4 超过 768 MiB，保留失败；仅释放排版缓存的 `20261102-201021-797c1b2f` 仍超限。最后在 PDF 生成/导出后释放可重建的 HarfBuzz 排版与解码图片缓存，保留原字体 face 和冻结 PrintJob；打印栅格可重新建立缓存，定向 6 项通过（`a4-print-cache.trx`）。Avalonia 宿主设置 `System.GC.ConserveMemory=5`，促使大型临时字体/解析缓冲更及时回收；没有手动强制 GC、减少受测进程或提高阈值。配置依据：[.NET GC 配置](https://learn.microsoft.com/dotnet/core/runtime-config/garbage-collector#conserve-memory)。

这一最终配置连续 PDF 两轮通过：`20261102-201138-5de00df5`、`20261102-201225-208db87c`。普通 E4 峰值分别 **791,449,600 / 782,094,336 字节**（约 755/746 MiB）；后一轮 100 页 **788,811,776**、2,000 页 **894,058,496**、高图像 **1,155,538,944** 字节。各场景首张画布 p95 **381.58 / 285.10 / 273.42 / 289.49 / 1093.98 / 179.62 ms**，图片缩放 **49.2 ms**。这仍是当前机器的工程结果，普通 PDF 内存余量需要在设备矩阵中关注。

最终 GC 配置下并发 `20261102-201254-ea03b14d` 通过，网格普通/样式/拖选 **9.4139 / 9.8408 / 18.0095 ms**，补全 **2.4 ms**、PDF 缩放 **48.8 ms**；Monaco `20261102-201306-067d6b3e` 输入、补全、开关保留与进程树上限通过。Linux 两轮证据对应此前相同 Web 补丁版本，最后的缓存释放/GC 配置尚未在 Linux 重跑，留给平台阶段复验。
最终 Windows 功能复验 `20261102-201442-14a442b6`：42 项通过，随后全解决方案构建 0 警告/0 错误。最终 Monaco 输入 p95 13.5 ms、补全 6.5 ms、打开 1137.16 ms，峰值 629,530,624 字节，十轮后增长 25,456,640 字节。

## 2026-11-03：修复离线资源副本不一致

PDF 预览报“Web 资源容量不匹配或超限”、脚本编辑器报“页面已停止”的本地复现来自同一资源目录：构建输出 `preview.html` 为 2745 字节，清单要求 2751 字节；`script.html` 为 677 字节，清单要求 680 字节。两个输出文件删过省略号，时间戳比正确的生成资源新，`PreserveNewest` 因而跳过覆盖；重新构建后错误仍然存在。

Avalonia 项目的 Web 资源在 build/publish 中改用 `IfDifferent`，修复更新但大小或时间戳不同的目标副本。运行时仍严格检查容量与 SHA-256；大小不匹配现在报告资源名、预期/实际字节数和重新构建/安装提示，不再与容量超限混为一谈。应修改 `QuickExcel.Web/src` 后生成完整资源与清单，不应直接修改输出目录中的 Web 文件。

- `scripts/test-avalonia-web-assets.ps1` 在独立 artifacts 目录实际构建/发布，改写两份 HTML、等长 JavaScript 和清单并设置未来时间戳，再增量构建/发布；两项修复检查通过，证据 `artifacts/avalonia/web-asset-tests/20261103-014908-6327a2b4`。
- 本地资源/草稿协议 **18 项通过**（包含等长修改仍被 SHA-256 拒绝），`artifacts/avalonia/web-resource-fix/6b208bda62f04bb38e18fd2435d4bd4a/web-resource-fix.trx`；修复前的真实宿主失败保留在同目录 `before/result.json`。
- 真实 WebView2 页面 **42 项通过**，覆盖 Monaco worker/编辑/保存/恢复和 PDF.js 预览/搜索/导出，证据 `artifacts/avalonia/web-tests/20261103-014946-27126167`。只运行相关专项，未运行全量回归或性能基准。
- 正式包 `artifacts/avalonia/releases/20261103-015012-6f4cdc85/application` 已通过生产自检并激活；开发输出已恢复正式资源。实机输入/打印等延期范围不变。

## 2026-11-04：脚本分类导航与创建时属性

脚本工作区左侧将主入口 `main.js` / `main.ts` 固定在顶部，下方为 UDF、Actions、辅助三个可独立折叠的分类，显示各组模块数量。主入口不重复进入分类；其余模块按既有绑定归类，双角色模块同时出现在 UDF 和 Actions 中。分类内按逻辑 ID 排序，折叠和展开不修改文档，普通草稿/状态刷新保留折叠状态。选择文件仍先确认草稿；诊断定位同步导航选中项，并展开目标分类。

“添加模块”在同一个对话框中定义逻辑 ID、角色和 JavaScript / TypeScript，创建对应扩展名和 Monaco 模型。角色和语言在创建后固定：移除原工具栏的“模块角色”和“JS / TS”修改入口，当前属性通过状态栏和文件提示只读展示。辅助模块不新增执行入口；UDF、Actions 或双角色在创建时一次性建立相应入口/绑定。已有模块属性和旧包/API v1 读取语义保留，包级替换与签名验证合同不变。

新建仍通过 `UpdateScriptModulesCommand` 提交，撤销旧授权和签名；取消、重复 ID、非法 ID 和大小写路径碰撞不会创建模块。没有改变隔离执行、编译、信任及只读保护规则。

验证：Avalonia Release 构建 0 警告、0 错误；`scripts/test-avalonia-web.ps1 -ScriptsOnly` 的 123 项脚本专项通过，证据 `artifacts/avalonia/web-tests/20261104-023935-d4d0d227`，包含分类展开/折叠截图、实际 Monaco 切换、原生创建对话框校验、XLSQ 属性往返和授权/签名撤销。首次检查误将新工作簿主入口当作双角色，已按其实际 UDF 默认绑定修正测试；失败证据保留在 `20261104-023725-f78fb996`。未运行全量回归或性能基准。

## 2026-11-04：角色示例、定义导航与紧凑文件列表

新建 UDF / Actions 模块分别使用角色模板，双角色包含两段示例。UDF 示例注册 `JS.<模块名>.DOUBLE`，演示参数校验和数值返回；无效参数抛出明确错误，由既有运行规则报告 `#SCRIPT!`。Actions 示例向当前单元格写入 42 并设置填充色，整体提交且可一次撤销。JS 使用 JSDoc，TS 使用显式类型或上下文推断；辅助模块仍提供可导入函数。模板名称兼容数字开头、大小写不同、带标点和较长的合法模块 ID，避免无效或重复注册。既有源码不自动替换。

诊断悬浮中的 `quickexcel.d.ts(行, 列)` 链接、Ctrl+单击和“转到定义”按钮接入 Monaco 定义提供器，通过宿主确认草稿后导航，并同步左侧选中项。F12 仍为控制台快捷键。左侧新增“定义文件（只读）”，可直接浏览 QuickExcel 声明；经跳转打开的内置 ECMAScript 声明也保留在此列表，支持编辑器内查找。定义模型与源码草稿分开，不进入版本确认、源码预算、保存或执行；只有当前工作簿模块与内置声明可打开，不访问外部 URI、系统文件或网络。QuickExcel 声明注释统一为英文。

所有文件条目改为 28 DIP 高的单行，选中背景铺满面板宽度，移除分类内容的横向留白。逻辑 ID、完整路径和属性保留在提示中；主入口置顶、分类折叠和源码的创建时属性规则保留。

验证：`ScriptModuleTemplateTests` 7 项通过，覆盖 JS/TS 的 UDF、Actions、双角色、错误参数、动作撤销及特殊模块 ID；`scripts/test-avalonia-web.ps1 -ScriptsOnly` 140 项通过，证据 `artifacts/avalonia/web-tests/20261104-031247-e73fa218`，包含真实诊断链接点击、精确行列定位、标准库/跨模块跳转、定义只读及源码保留，紧凑列表截图已复核。仅运行相关专项，未运行全量回归。

### 纯 JavaScript 工作簿首次打开定义文件

Monaco 按语言延迟注册服务。纯 JavaScript 页面首次打开 `quickexcel.d.ts` 时，定义浏览代码在创建 TypeScript 模型之前调用 `getTypeScriptWorker()`，导致 `TypeScript not registered!` 并触发既有页面失败保护。原先定义导航检查在 TypeScript 工作簿中进行，未覆盖此初始化顺序。

QuickExcel 声明现在直接使用已打包的文本；内置标准库声明通过当前源码模块已初始化的语言 worker 获取。定义仍为只读模型，保留原有草稿确认、模块边界和导航规则。新增两个独立冷页面检查，分别从纯 JavaScript 工作簿首次打开 API 声明和 `lib.es5.d.ts`，并验证返回后源码、编辑状态及语言服务正常。

修复前实际 WebView2 检查复现相同错误，证据 `artifacts/avalonia/web-tests/20261104-033141-1d62ad58`；修复后 `scripts/test-avalonia-web.ps1 -ScriptsOnly` **146 项通过**，证据 `artifacts/avalonia/web-tests/20261104-033213-a4cd52f2`。Avalonia Release 构建 0 警告、0 错误；未运行全量回归或性能基准。

`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 已生成并激活 `artifacts/avalonia/releases/20261104-033257-b8ec970e/application`，8 项生产自检通过。活动清单身份及正式 Web 资源一致性已核验，开发构建输出也已恢复正式 Web 资源。

## 2026-11-04：默认 TypeScript 与主入口角色

桌面新建空白工作簿使用 API v2 的 `scripts/main.ts`，新建模块对话框默认选择 TypeScript，仍可显式选择 JavaScript。主入口在文件提示和状态栏统一显示“主入口 · TypeScript”（已有 JavaScript 主入口显示相应语言），与顶部固定文件导航使用相同的主入口判定。

新建主入口同时声明 `formula-functions` 和 `actions` 绑定，可注册 UDF、操作和事件，并支持现有 `function main(workbook: ExcelScript.Workbook)` 操作入口。主入口不重复归入 UDF / Actions 分类，未新增文件格式中的绑定种类。打开已有工作簿不改写语言、源码、绑定或签名，API v1 和示例工作簿保留既有执行语义。

验证：`WorkbookWorkspaceTests` 与 `ScriptModuleTemplateTests` 共 **17 项通过**，包括新建/保存重开默认值、旧 JS 保留、TS 主入口内 UDF 计算与 `main()` 操作及撤销；`scripts/test-avalonia-web.ps1 -ScriptsOnly` **148 项通过**，证据 `artifacts/avalonia/web-tests/20261104-035517-3d9e5735`，验证实际 Monaco 语言、主入口属性、创建对话框默认值与分类导航。首次 Web 打包因临时目录重命名失败（EPERM）中止，重试成功。导航截图已复核；测试后恢复正式 Web 资源并完成 Avalonia Release 构建，0 警告、0 错误。未运行全量回归或性能基准。
