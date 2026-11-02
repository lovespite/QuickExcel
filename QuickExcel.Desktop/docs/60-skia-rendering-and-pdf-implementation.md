# Skia 共享渲染与直接 PDF 实施记录

2026-11-02，接续 [59 号方案](59-skia-rendering-and-pdf-plan.md)。**P0–P4 声明范围完成**：共享渲染、直接 PDF、页面预览、主网格和旧依赖清理已接入，Windows/Linux 实际输出、相关功能及原有滚动预算通过。最终候选包为 `artifacts/publish/20261102-130713-e1444da9`，同步至 `artifacts/QuickExcel`。本记录保留中间失败样本和测量修正依据；物理 IME、跨显示器、真实打印设备、整应用跨平台及既有 M4/M6 门禁不在本次完成声明内。

## 1. 模块与生产路径

| 模块 | 实际职责与边界 |
| --- | --- |
| `QuickExcel.Rendering` | `net10.0`，只引用 Core；DIP 几何、字体/文字/单元格请求、预算和不可变 PrintJob |
| `QuickExcel.Rendering.Skia` | `net10.0`；字体回退、HarfBuzz/RichTextKit 排版、单元格、图片、基础图表、页面 PNG 与直接 PDF |
| `QuickExcel.RichTextKit` | 仓库内固定版本的 Apache-2.0 排版源码；来源、原始哈希及本地改动见 [SOURCES](../../third_party/RichTextKit/SOURCES.md) |
| Desktop / Grid.Wpf | WPF DPI、冻结/滚动几何、只读位图上屏、选区、原生编辑器/IME 和系统打印适配 |
| Rendering.Test / Rendering.Smoke | 无 WPF、Infrastructure 或 worker 引用的测试与独立可执行程序；固定 Core 快照，不加载脚本 |

Core 不引用 Skia；条件格式由 Desktop 组合根提供显示样式委托。绘制不重算、不启动脚本、不改变文档修订/dirty/历史。打印入口仍先提交草稿并完成既有全量重算，受阻公式继续禁止打印，不将来源缓存标为新结果。链接保持惰性；富文本替换仍需显式转换。

网格普通/样式格均走 `SkiaRenderer.DrawCell`，自动尺寸走同一文字布局。可见格的 Skia picture 与只读 WPF raster 分开缓存；同一可见集合内，冻结区/滚动区复用 DrawingGroup 并平移，超大格只栅格化视口裁剪。快照、尺寸、DPI、缩放或可见集合变化使相应缓存失效；卸载释放字体、picture、图片和缓存。WPF 编辑 TextBox 继续使用宿主字体 API，不属于内容渲染器。

删除原 GDI+/FormattedText 字形生成、WriteableBitmapEx 绘制实现和 `msvcrt.dll` 导入，移除 Grid 的 `UseWindowsForms`。保留实际复用的几何/字体描述和上游许可证。Windows 上屏适配使用私有锁定像素缓冲与冻结副本，不复写合成器正在使用的 bitmap。

专项验证另外发现并修复两处既有 UI 问题：保存时不再为未改变的脚本提交命令，从而允许只读公式工作簿原包保存；公式栏统一合并只读公式、富文本及保护状态，避免一次刷新覆盖另一种保护。

## 2. 文字、图片与 PDF

文字覆盖中英文、组合字符、补充平面字符、阿拉伯文/希伯来文混排、普通及富文本字体继承、颜色/装饰、手动换行、自动换行、缩进、旋转、缩小填充和数字格式填充字符。测试固定八个 Noto 字体并记录 SHA-256；产品使用安装字体，替代/缺字产生有位置、数量有界的诊断。不同机器安装字体不同，不承诺同像素布局。

PNG/JPEG/GIF/BMP 由 Skia 解码；TIFF 使用 LibTiff.NET 的第一页适配。五种格式在 Windows 和实际 Linux 进程均验证，GIF/TIFF 的多帧默认使用第一帧/页。透明像素、黑白投影、预算与原字节不变有专项断言。图片原始字节仍由存储层保存。图表保持簇状柱形、基础折线与饼图原有支持范围。

页面复用 Core 分页、重复标题、隐藏/筛选、分离区域、手动分页和打印缩放。常规文字是嵌入子集字体的 PDF 文字对象，边框/图表保持矢量；HarfBuzz 输出同时保留原 UTF-8 与 cluster 映射，独立解析器检查最终 PDF 的实际文字。坐标使用 DIP，PDF 统一乘 72/96；Skia PDF RasterDpi=600 避免 MediaBox 取整超出 0.2 pt 验收阈值。

直接 PDF 不使用 XPS、打印队列或 Microsoft Print to PDF。输出在目标同目录写唯一临时文件，关闭/flush 后用 PdfPig 验证页数、纸张及页面内容可解析，再检查取消并替换。`PdfOutputStream` 逐次写入限制输出字节，并捕获原生回调内的托管写入/flush 异常，在文档调用返回后重新抛出。取消、预算失败、替换失败和关闭预览保留已有目标，清理本次临时文件。Windows/Linux 均实际验证成功替换与失败保护；Linux 在独立 64 KiB tmpfs 上实际触发 ENOSPC，旧 PDF 保留、临时文件清理，进程正常返回。

预览按可见页异步加载，缩放请求合并 50 ms、旧请求取消、代次校验后上屏；页图和缩略图共用任务与缓存。打印内容提示在预览前统一展示，省略只作用于本次任务。PDF 与预览共用异步渲染门禁，关闭时取消并异步等待，不在 UI 线程等待后台原生绘制锁。

系统物理打印保留 Windows `PrintDialog`，使用同一任务的最高 300 DPI 有界栅格适配；直接 PDF 是文字/矢量路径。没有宣称本次完成真实打印机、跨显示器输入、物理 IME、PDF/A、PDF/UA、PDF 阅读器或整应用跨平台验收。

## 3. 资源预算与取消

| 资源 | 当前限制 |
| --- | --- |
| 整任务分页 | 保留 Core 的 2,000 页、2,000,000 格限制 |
| 单页 | 100,000 格、2,000,000 个文字 UTF-16 单位 |
| PDF 文件 | 128 MiB；每次原生写入前检查，逐页及关闭后传回错误 |
| 单张解码图片 | 16,000,000 像素；解码前检查尺寸 |
| 图片缓存 | 每 renderer 64 MiB |
| 预览页图缓存 | 最多 8 张且合计 64 MiB；按预算限制 DPI，最高 600 |
| 网格 picture | 4,096 项且估算 32 MiB；小格 raster 合计 64 MiB |
| 字体 | 128 个 face；请求/coverage 查找另有数量上限 |
| 诊断 | 每字体解析器最多 1,024 项 |

这些是各缓存/对象的限制，不是进程总私有内存上限；WPF 还持有当前可见帧，生成新帧也有短时重叠。独立 smoke 和基准记录进程私有内存。单元格、图片/图表和 PDF 页边界检查取消；原生解码、塑形或 PDF 单次调用不能中断，因此不承诺固定毫秒取消上界。单页和解码预算限制一次工作的输入量，旧 XPS 的 60 秒驱动等待不再套用。

## 4. 固定依赖与平台

生产版本：SkiaSharp/SkiaSharp.HarfBuzz/Skia native assets **4.154.0**，HarfBuzzSharp native assets **14.2.1.400**，LibTiff.NET **2.4.660**，PdfPig **0.1.16**，RichTextKit **0.6.170** 固定源码。NuGet 内容哈希保存在 lock 文件，完整原生第三方 notices、MIT/Apache/BSD 许可证随包复制。详见 [依赖来源](../../QuickExcel.Rendering.Skia/SOURCES.md)。

Windows x64：Windows build 22621、.NET 10.0.12，CPU i9-13980HX、32 GiB。Linux x64：实际 Ubuntu 24.04.5 虚拟机，kernel 6.8.0-147、glibc/fontconfig/freetype/ICU 运行库；无显示服务器或 Windows Desktop Runtime。使用自包含 linux-x64 smoke 的真实 ELF 原生库完成输出，`ldd` 和运行环境保存在 `artifacts/skia-final/linux-io-run.log`。此证据不扩展为所有 Linux 发行版或 macOS 支持声明。

固定字体为测试用 OFL 资产；MuPDF/PyMuPDF 1.26.7 是独立验证工具；QEMU/Ubuntu 与 Python 依赖均仅在 ignored artifacts，不随产品或 worker 发布。发布脚本检查所需渲染程序集、原生 DLL 和许可证，同时禁止它们混入 ScriptWorker 载荷。

新增 [Windows/Linux CI](../../.github/workflows/rendering.yml) 运行固定字体、共享测试、smoke 和独立 PDF 验证；**远端 CI 尚未运行**。本地 Linux 已实际运行独立 smoke，未把 Windows xUnit 结果称作 Linux xUnit 结果。

## 5. 定向验证与原始证据

所有证据位于本仓库 ignored `artifacts`，没有执行全量 Core/桌面回归。

| 验证 | 结果与位置 |
| --- | --- |
| Release 解决方案构建 | 0 错误、0 警告，`artifacts/build` |
| 共享渲染 xUnit | 20/20；`artifacts/skia-final/tests/skia-io-final.trx`；包括绘制中取消、原文件保护、图片、native 预算、写入/flush 错误及替换失败 |
| Core 分页/E4 | 31 项定向测试通过；网格几何 13 项通过，`artifacts/skia-final/tests/grid.trx` |
| Windows/Linux 无界面 | 最终 `artifacts/skia-final/windows-io`、`artifacts/skia-final/linux-io`；五种图片、Unicode、原子替换、取消和预算失败保护均通过；原先 windows/linux 目录也保留 |
| 实际存储/缺库失败 | `artifacts/skia-final/linux-io-run.log`；ENOSPC 原文件保护与临时清理通过；独立载荷移走 libSkiaSharp.so 后明确加载失败，无后备路径/PDF 输出，完整错误见 `missing-native-linux.log` |
| 独立最终 PDF | 上述目录的 `independent-validation.json`；MuPDF 实际解析、字体嵌入、文字、页归属、纸张、矢量和逐页局部像素检查通过 |
| XLSX 渲染/保存/富文本/附件/剪贴板/填充 | 148 项通过；`artifacts/skia-final/xlsx-rendering-7` |
| 网格集成 | 最终候选包 `artifacts/validation/m5-grid/20261102-130223-96b325d5`；含实际 WPF 生命周期、冻结/合并/宽列/缩放、编辑所有权及仅内容层失效时的完成通知 |
| 工作区 / 输入 | 157 / 95 项通过；`20261102-120249-bbfd4ca6` / `20261102-120307-c69a74da` |
| 格式效率/自动尺寸 | 24 项通过；`20261102-113635-4ce05338` |
| 最新报表 / 预览关闭 | 最终候选包 72 项通过；`20261102-130243-16cf45f2/Reporting`，含队列取消、缓存释放与旧 PDF 保留 |
| E4 独立视觉/Excel 页归属 | 同目录最终五页 PDF 全部通过；中文可提取、嵌入 Type0 字体、局部误差最大 0.08490；Record 19/20 分页与既有 Excel 输出一致 |
| 最终候选发布包 | `artifacts/publish/20261102-130713-e1444da9`；记录 11 项渲染依赖版本/内容哈希、两项 native DLL 哈希及 FastGridSurface。原生库/许可证存在，来源与第三方总清单随包复制，worker 无渲染载荷；既有 QuickJS 10,000 次调用 smoke 通过 |
| 候选包复验 | 网格、报表和横向预算直接测 `20261102-130131-7cac08d4`，五页最终 PDF 独立检查通过，见 `artifacts/skia-final/package-io-pdf-validation.log`。最终包补齐发布文档/来源清单，全部 **32 个 DLL/EXE SHA-256 与该受测包相同**。构建使用 `-SkipTests`，元数据如实记录未执行全量测试，专项证据按本表关联 |

独立视觉检查按 128×128 占用像素块比较，允许 1 像素抗锯齿邻域与 64 通道差异，各块错配率须小于 0.10；不是以全页空白稀释差异。另对文字、页归属和纸张作硬断言。相同固定字体跨 Windows/Linux 比较分页完全一致、实际提取文字一致、指定文字几何误差小于 0.5 pt。脚本为 [verify-skia-output.py](../../scripts/verify-skia-output.py) 和 [verify-e4-visuals.py](../../scripts/verify-e4-visuals.py)。

## 6. 性能与测量修正

16.67 ms 横向输入到 WPF CPU 绘制完成 p95、33.34 ms 拖选自动滚动间隔 p95 预算保持不变，最终均通过。它们不是物理显示延迟。基准结果以实际预算返回 PASS/FAIL，失败时脚本非零退出并保留环境/原始样本。

排查时新增网格、内容和选区绘制时间戳以及 .NET 运行栈采样。`20261102-123820-24d8a422` 的实际绘制结束 p95 约 10–12 ms，但 ApplicationIdle/ContextIdle 完成观察额外等待可达约 20 ms；慢样本在内容和选区完成后仍滞留。现改为 Render 优先级观察，同时把宿主内容层的 pending 状态纳入 `IsRenderPending`，从网格、选区及宿主完成回调统一判断。不会提前跳过内容绘制，也不再等待无关 Win32 输入让 Dispatcher 空闲。“仅宿主内容失效”的新回归确认完成前仍为 pending、完成后确实发生绘制。完整时间戳在横向 raw phases，采样在 `artifacts/skia-profile`。

旧引擎从 HEAD `35877e7115c854f01929380e1e33a2a5720cf1e4` 在 artifacts 独立重建，仅应用相同完成观察/pending 协议，保留原 GDI/WPF 绘制。新旧实现均用修正后的协议重跑；生产横向输入仍按合成帧合并，没有改为每次输入立即绘制。

| 样本 | 普通格 p95 | 样式格 p95 | 说明 |
| --- | --- | --- | --- |
| 初始 P0 `20261102-101939-aa074f6a` | 9.35 ms | 16.10 ms | 原测量器，当时通过 |
| 同机重建旧引擎 `20261102-115027-0e930acc` | 30.17 ms | 44.15 ms | 旧 idle 完成观察；失败样本保留 |
| Skia pane 缓存 `20261102-115339-b5b9d6f0` | 28.01 ms | 30.09 ms | 超过预算；不能以优于重跑基线替代通过 |
| Skia `20261102-122221-a810b1a2` | 33.38 ms | 24.67 ms | 旧 idle 完成观察；失败样本保留 |
| 修正完成观察后的旧引擎 `20261102-125256-b80d14a8` | 10.61 ms | 12.96 ms | 全场景通过；原有绘制实现 |
| 修正完成观察后的 Skia `20261102-124504-76090eb4` | 15.05 ms | 14.09 ms | 全场景通过 |
| 最终候选包 Skia `20261102-130205-b10c0bb1` | **15.36 ms** | **11.68 ms** | 全场景通过；输入至视口应用 p95 0.90/0.57 ms |

最终 Skia 每场景 65 次视口通知仅新增 525 格内容；普通/样式格分配约 21.4/26.6 MB；缓存原生 picture/像素约 7.3 MB，场景末进程私有内存约 237/247 MB，隐藏内容场景末回落至约 246 MB。隐藏内容层 p95 为 9.05 ms。相同观察方式的旧引擎样式格分配约 60.1 MB，但普通格尾延迟仍优于当前 Skia；本次不宣称所有性能指标都提升。

拖选已修复并通过：原 16 ms timer 与 Rendering 共享墙钟门禁，timer 可能先更新时钟，导致紧邻的合法合成帧被跳过；现在 Rendering 只按帧编号去重，50 ms timer 仅作遮挡备用。最终网格源码 `20261102-125335-f0db9af6` 全部拖选预算通过，自动滚动 120 次更新、间隔 p95 **18.87 ms**，低于 33.34 ms，纯选区拖动仍不读取模型。此前 `20261102-121445-95a2c0b0` 的 35.70 ms 失败样本及修复后 `20261102-122047-143aa530` 的 18.86 ms 样本保留。

`FastGridBenchmark` 的导航、巨型选区、筛选及对应预算通过，见 `20261102-111601-bcea4fcc`。真实不等宽列 `D:\Downloads\笼卡demo.xlsx` 的顺序独立滚动压力通过终点/积压和原文件不变断言，最终观察协议证据 `20261102-125422-bad4e503`；SHA-256 `21E8714CBB95CA1105B50E699E21246CC364E6B00147C0F1B614BFEB60F518F6`。该工作簿仅只读加载，没有改写。

全视口光栅、不同等待器/调度、DrawingGroup 包装等中间实验与失败样本均保留。重复帧计数实验未发现横向回调被去重丢弃；当前帧 DrawingVisual 提交、100 ms 保持合成订阅、WPF DrawingBrush GPU 缓存均未改善尾延迟，已撤回。软件合成对照 `20261102-121138-b07f259b` 也未通过预算，生产默认仍为自动硬件选择。`20261102-111650-ee41c5a7` 的拖选样本意外与源文件压力重叠，不能作为隔离性能结论；最新样本均顺序运行。没有为了满足门槛上调预算或选取最好的一次。

P3 本地定向功能及性能门槛已关闭。真实跨屏、物理 IME、系统打印设备及现有 M4/M6 门禁仍单独开放；它们与共享模块在 Linux 无界面运行的证据不是同一验收。

## 7. 阶段验收

| 阶段 | 完成依据 |
| --- | --- |
| P0 | 固定包/字体/许可，原引擎基线与资源预算，Windows/Linux 实际 native 加载及五种图片验证 |
| P1 | 无 Desktop/Infrastructure/worker 引用的 net10.0 模块；两平台 PNG/PDF、Unicode、富文本、物理单位、分页及独立对照通过 |
| P2 | 共享不可变任务接入报表、按需预览和直接 PDF；E4 72 项、最终五页 PDF/Excel 页归属、取消/关闭/预算/真实磁盘失败保护通过 |
| P3 | 普通/样式格、附件、自动尺寸共用布局；缩放/DPI/冻结/隐藏/合并/编辑相关验证及横向、拖选、真实宽列压力通过；缓存与生命周期有界 |
| P4 | 旧字形/GDI/CRT/Windows Forms 依赖移除；来源和许可证保留；发布依赖、原生库、worker 排除及候选包专项通过 |

远端 CI 尚未执行，macOS、完整应用非 Windows 运行、物理设备及既有 M4/M6 门禁均保留为独立后续项。提交/发布到远端不属于本次工作；工作区修改尚未提交。

## 8. 重现命令

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Rendering.Test\QuickExcel.Rendering.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter 'FullyQualifiedName~E4ReportingTests|FullyQualifiedName~WorksheetPrintingTests' -v minimal
.\scripts\get-rendering-test-fonts.ps1
dotnet run --project .\QuickExcel.Rendering.Smoke\QuickExcel.Rendering.Smoke.csproj -c Release --artifacts-path .\artifacts\build -- artifacts\rendering-output artifacts\skia-p0\fonts
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Reporting
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode GridIntegration
# 基准单独顺序执行，并检查 JSON 的 withinBudget / passed。
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode HorizontalScrollBenchmark
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode FastGridDragBenchmark
.\scripts\build-quickexcel.ps1 -SkipTests
```

独立验证需要测试环境的 PyMuPDF/Pillow/numpy；脚本可指定仓库 artifacts 中的依赖目录。Linux smoke 发布同样必须使用仓库内 `--artifacts-path`，将自包含输出和固定测试字体复制至实际 Linux 执行后取回证据；仅交叉 publish 不算运行通过。
