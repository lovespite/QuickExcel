# 主界面网格迁移（2026-10-29）

主工作区直接使用 `QuickExcel.Grid.Wpf` 的 `FastGridControl`，由 `FastGridSurface` 适配。旧 `WpfGridSurface`、DataGrid 单元格/行容器、冻结副本网格、旧尺寸事件和 Adorner 实现、旧 Grid.xaml 主题以及 `FastGridWindow` 试验入口已移除。此前的 [23 号记录](23-fast-grid-implementation.md) 保留历史对照证据。

## 当前状态与缺陷关闭（2026-10-30）

**横向拖动导致整桌面停顿：已完成。** 用户于 2026-10-30 补充确认该问题已经解决，并要求更新遗漏的进度记录。以此次用户确认关闭该缺陷，不再将其列为后续阻塞项。

当前代码默认采用 `RenderMode.Default`；`--software-rendering` 只保留为诊断选项。下文的软件渲染尝试、撤回和线程/消息路径分析是关闭前的历史记录，其中“根因未确认”“尚不能证明修复”等表述反映当时证据，不代表缺陷仍开放。本次不新增根因结论，也不声称重新采集了线程栈、GPU 跟踪或性能样本。

主网格迁移、筛选滚动范围、行列尺寸拖动/双击自动调整、合并选区及显示修复已有下文专项证据。新主网格真实 OS IME、跨屏 DPI、完整 UI Automation/读屏以及 M5/M6 的完整质量门禁仍独立待验收。项目状态与后续顺序统一见 [实施计划](05-implementation-plan.md)。本次为文档同步，没有运行新的构建、测试或发布。

## 缩放清晰度与格式层对齐（2026-11-02）

用户报告：非 100% 缩放时普通单元格文字、列头、行头模糊；填充颜色或改用非常规格式后，单元格边框相对位图内容错位 1~2px，文字同时下沉。

根因有四处，均在渲染路径而不在文档或命令：

1. `FastGridControl` 的内容位图只按窗口 DPI（`DpiDetector`，主屏系统 DPI）分配像素，而行列几何按 DIP 存储、由工作表缩放以 `LayoutTransform` 施加在 `SheetPageContent` 上。缩放不为 100% 时位图被合成器整体放大，所有栅格文字（普通单元格、列头、行头）按小数像素重采样变模糊；格式层用矢量 `FormattedText` 绘制，因此只有已格式化单元格保持清晰。跨显示器 DPI 不一致时同样退化为放大位图。
2. 位图以 `(int)ActualWidth - 2` 的整数 DIP 宽度显示，而缓冲区像素宽度是 `ceil(width × 比例)`；两者不等时 Image 会把已按正确比例栅格化的内容再缩放不到一个像素，进一步打乱字形列。
3. 格式层按 DIP 画整格矩形并用 1 DIP 的画笔沿边线描边；位图按设备像素绘制、且单元格矩形包含尾部分隔像素（`IntRect.Right = Left + size`，填充比矩形内缩一像素）。两者坐标系不同，缩放非整数时边框被抗锯齿摊到 2 像素并整体偏移；文字则用 `FormattedText.Height` 直接居中，与位图“向上取整行高后居中”的度量不同，出现 1px 下沉。格式层在未指定字号时还回退 11pt（14.67 DIP），而位图使用网格字号 14 DIP。
4. 更明显的“下沉”来自垂直对齐：样式默认 `Vertical = 2`（底端，与 Excel 一致，功能区也按此勾选），格式层按底端绘制，而位图层（`RenderBlock`）此前无视样式始终居中。单元格一旦套用填充/字体等样式就改由格式层绘制，文字立即从居中跳到行底。

处置：`FastGridControl` 新增 `RenderScaleX/Y`，由控件自身到组合器的变换（祖先布局缩放 × 监控器 DPI）测得，替代项目内的 `DpiDetector`：位图按该比例分配像素并以自然 DIP 尺寸显示，字形按同一比例栅格化，行高列宽、命中测试、选区/编辑框/格式层片段全部换算回 DIP；缩放或 DPI 变化、设备重新连接时重建位图与缓存。缓存键加入比例，保留原有视口淘汰与图层分工（内容位图负责普通文字与底色，格式层负责样式化单元格）。`FastGridArtwork` 的填充、网格线、样式边框改为整设备像素的矩形条带：填充从首像素之后开始、止于尾分隔像素之前，四条边线落在与位图相同的像素行/列；文字位置沿用位图的 `ceil(行高×比例)` 与整数取整居中，左右内边距、可用宽度按比例换算，未指定字号时使用 `CellFontSize`。位图渲染同时支持单元格垂直对齐（`FastGridCellImpl.VerticalAlignment`），使未格式化与已格式化单元格的垂直定位完全一致；由于工作表默认对齐为底端，普通数据文字会比此前居中时更靠下，行列标题仍居中。滚动条外边距同步改为 DIP（此前在高 DPI 下按像素值设置）。

限制：本节是代码修复记录。当前执行环境 shell 不可用，未运行 Release 构建、`GridIntegration`/`Workspace`/`Interaction` 专项或真实缩放屏幕检查；`RenderScale` 依赖可视树变换，首次布局前按窗口 DPI 回退。真实 IME、跨屏 DPI 与显示帧率门禁不变。

### 格式层仍偏移约 2px 的根因（2026-11-02 补记）

上一节的整设备像素几何本身是对的，偏移来自更外层：`RebuildViewport()` 给位图 `Image` 显式设置了 `Width`/`Height`（`PixelWidth / RenderScale`，约等于 `(int)ActualWidth - 2`），而对齐方式仍是默认的 `Stretch`。WPF 在“显式尺寸 + Stretch”时按 `Center` 排列子元素，于是整张位图在 `imageGrid` 中被居中，原点偏离控件原点约 1 DIP（200% 缩放下约 2 设备像素）。`FastGridArtwork`、`GridSelectionOverlay`、内联编辑框和 `HitTestModel` 都从控件原点 (0,0) 起算，因此只有位图整体平移：格式单元格的填充/边线与相邻栅格单元格错开 1~2px，右侧还会露出位图层未被格式层覆盖的底色细条。

处置：`Image` 固定 `HorizontalAlignment="Left" VerticalAlignment="Top"`，位图原点与矢量层共用控件原点；新增 `FastGridControl.RasterOrigin`（DIP）并在 `CellLayoutSmoke` 中断言其为 (0,0)，避免以后再次因显式尺寸被布局居中。

验证限制：本轮同样只有静态核对，未运行 Release 构建与 `CellLayout` 冒烟；需在真实缩放（100%/125%/150%/200%）下目视确认。


## 行为与边界

- 文档仍由 `EditorController` / `DocumentSession` 写入。公式栏和原生 TextBox 保留原地址草稿、组合输入门控、显式公式粘贴和原子撤销。忙碌时被窗口拦截的组合事件不能激活网格草稿。
- 主界面使用已有导航、剪贴板、保存、查找、结构和 Ribbon 命令。绑定新文档重置地址；普通计算和格式刷新保留选区、滚动位置、模型和输入控件。过滤掉当前行时定位可见标题。
- 冻结行列由同一控件绘制。筛选按钮、编辑层和图片/合并层按冻结及滚动分区裁剪。字体、字号、对齐、换行和各边框由 `FastGridArtwork` 投影到内容层，选区在其上独立绘制。图片解码继续限制像素数和解码宽度。
- 行高和列宽拖动由主窗口处理，拖动时只更新视图，结束时提交一个尺寸命令，取消不产生历史。尺寸依旧以行高 DIP / 列宽字符单位保存，几何会按设备像素取整。
- 隐藏当前行列后选择最近可见位置；“取消隐藏”包含当前选区两侧相邻的手动隐藏行列。区域操作可包含手动隐藏记录，筛选造成的不连续选区仍拒绝；复制跳过手动隐藏和筛选隐藏行，粘贴仍不能跨筛选隐藏行。
- `FastGridBenchmark` 只测当前实现；旧 DataGrid 对照代码已删除，历史原始样本保留在 artifacts 和 23 号记录中。没有修改 ExcelIO 源码、Core 模型或脚本隔离策略。

## 验证与修复过程（历史记录）

以下按原执行轮次保留证据；横向整桌面停顿的最新状态以上方关闭记录为准。

### 恢复默认自动渲染（2026-10-30）

用户反馈软件渲染虽然不再卡死其他窗口，但绘制速度慢到无法正常使用，因此撤回上一轮默认启用软件渲染的处置。普通启动恢复 `RenderMode.Default`，由 WPF 自动选择渲染路径；只有显式传入 `--software-rendering` 才使用 `SoftwareOnly`。`--default-rendering` 对照入口继续兼容。验证脚本的默认期望同步恢复为 `Default` / `automatic-default`。

软件渲染仅保留为诊断选项，不能作为此问题的可用修复。此前硬件合成方向的排查线索及原始测试证据保留，整桌面停顿的根因仍未确认。

撤回后 Desktop Release 构建通过（仅既有 NU1900 警告），无渲染参数的 GridIntegration 通过，实际配置确认为 `Default` / `automatic-default`。证据：`artifacts/validation/m5-grid/20261030-023425-a9ad49d9/GridIntegration/`。未运行完整回归或重复性能基准。

### 默认启用软件渲染的尝试（历史处置，已撤回，2026-10-30）

用户进一步确认：同一操作只在默认渲染模式出现整桌面停顿，`Start-QuickExcel-Software.cmd` 的软件渲染模式不出现。这个人工对照将主要疑点指向 WPF 硬件合成/显示驱动路径，但没有故障时的原生线程栈，仍不能确定具体驱动调用或声称已找到 GPU 死锁根因。

复查 FastGrid 的事件与资源路径：

1. `hscroll_Scroll` 的 `ThumbTrack` 仅覆盖最新位置；`QueueHorizontalScroll` 每个控件最多订阅一次 Rendering，另有一个后备定时器；处理后解除订阅。`CheckInvalidation` 只保留一个排队绘制请求。没有发现全局输入钩子、`BlockInput`、`AttachThreadInput` 或循环同步发送跨窗口消息的生产路径。
2. `HorizontalScrollRendering → RenderGrid → RenderRetainedContent` 每次提交新的、冻结的 `DrawingImage`；缓存未命中时栅格化可见 tile，然后交给 WPF 上屏。缓存按视口淘汰，私有 `_tileCanvas` / `_drawBuffer` 不直接交给合成器；不能仅因 `BitmapContext` 中有 `WriteableBitmap.Lock` 就断定正在等待 GPU。新位图/绘图资源的提交是硬件模式下需要进一步采样的边界。
3. 纠正一个容易误判的细节：`GrayScaleLetterGlyph.CreateGlyph` 的 `RenderTargetBitmap.Render` 确实是同步调用，但所查 .NET 10 WPF 基线源码中，`RenderTargetBitmap.FinalizeCreation` 传入 DEFAULT flags，`CMILFactory::CreateBitmapRenderTarget` 实际创建 `CSwRenderTargetBitmap`。所以它是冷字形生成的 UI 线程成本，不能直接等同于 GPU 渲染；这部分在两种模式下都存在。[RenderTargetBitmap 源码](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/PresentationCore/System/Windows/Media/Imaging/RenderTargetBitmap.cs)、[原生工厂源码](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/WpfGfx/core/api/api_factory.cpp)。
4. WPF `Thumb` 在 MouseUp 时释放捕获，之后才发 DragCompleted/EndScroll；FastGrid 没有在横向 Scroll 处理器重新捕获鼠标。UI 线程阻塞会延迟这个流程，但捕获本身不足以解释其他进程窗口失去响应。[Thumb 源码](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/PresentationFramework/System/Windows/Controls/Primitives/Thumb.cs)、[SetCapture 语义](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setcapture)。

处置采用已由用户确认有效的进程级规避：`App.OnStartup` 在创建任何窗口前默认设置 `RenderOptions.ProcessRenderMode = SoftwareOnly`。没有修改系统注册表或显示驱动，也没有改变 FastGrid 的输入/滚动语义。`--software-rendering` 继续兼容；`--default-rendering` 显式恢复 WPF 自动选择模式，供后续诊断使用，不保证一定使用 GPU；两者同时出现时软件模式优先。软件渲染可能增加 CPU 成本，这次不声称关闭性能门禁。[微软的进程级软件渲染说明](https://learn.microsoft.com/en-us/dotnet/api/system.windows.media.renderoptions.processrendermode?view=windowsdesktop-10.0)。

`test-m5-grid.ps1` 新增 `-DefaultRendering` 对照参数，并对每次运行的 `rendering.json` 同时核对实际进程配置和选择来源（`software-default` / `explicit-software` / `explicit-default`），避免无参数启动误测成硬件模式。旧的软件启动脚本仍可使用。

本轮 Desktop Release 构建成功，0 错误；有既有 NU1900 数据源警告及 apphost 文件占用后的复制重试。未终止用户进程。无渲染参数的 GridIntegration 通过，横向专项 31 项，`rendering.json` 确认 `software-default` / `SoftwareOnly`，证据：`artifacts/validation/m5-grid/20261030-022853-7b42a68c/GridIntegration/`。

单独运行原工作簿六场景压力检查，全部到达终点，原文件 SHA-256 保持 `21E8714CBB95CA1105B50E699E21246CC364E6B00147C0F1B614BFEB60F518F6`。两张表宽列内容提交间隔 p95 为 26.19 / 27.72 ms；首次全宽场景最大间隔仍有 523.18 ms，说明冷路径性能问题仍在。原始记录：`artifacts/validation/m5-grid/20261030-022909-f9bd4a5e/HorizontalScrollBenchmark/`。这些是 WPF 路由事件压力检查，不能替代物理鼠标或 GPU/DWM 跟踪；本轮未运行完整回归。

### 横向拖动输入停顿与消息路径审计（2026-10-30）

用户补充的准确现象是：连续拖动时鼠标仍能移动，但点击其他窗口无反应；松手后还要等待数秒或更久。此前应用内内容提交节拍的改善不能证明这一问题已经修复。

检查 Desktop/Grid 的生产代码及 .NET 10 WPF 源码后，未发现直接循环调用 `SendMessage` / `PostMessage`、全局鼠标钩子或 `AttachThreadInput` 的路径。`ThumbTrack` 只覆盖一个 `_pendingHorizontalScroll`，`CheckInvalidation` 用 `_renderQueued` 合并绘制请求；WPF Dispatcher 自己也合并唤醒消息。因此没有源码证据支持“每次 MouseMove 无界新增 Windows 消息”。需要区分队列长度与单次处理/等待耗时：有界队列仍可能长时间无法处理输入。

确认并移除一处多余排队：`HorizontalScrollRendering` 原先调用 `ApplyPendingHorizontalScroll` 时，视口失效会排入一个 Render 优先级回调；紧接着又同步 `RenderGrid`，使排入的回调通常只做空检查。本次在同帧处理外包裹已有 invalidation context，避免这次冗余排队。该缺陷不构成已证实的整桌面停顿根因。

其余重点是渲染/布局与同步等待，而非消息发送数量：

- `FastGridControl_RetainedContent` 每次内容提交都创建 `DrawingImage` 并设置 `Image.Source`；该依赖属性影响 Measure/Render。视口通知还更新编辑框/筛选按钮位置、裁剪及格式层。在 Rendering 回调内产生布局失效可能使 WPF 在同帧再做布局；显式 Image 尺寸限制了传播范围，尚未测得这部分实际耗时。
- `GrayScaleLetterGlyph.CreateGlyph` 在缓存未命中时同步调用 `RenderTargetBitmap.Render`；WPF 的 `BitmapVisualManager` 使用同步通道，存在 UI 线程等待的可能。字形命中后不重复调用，且样式文字已由格式层负责，不能把所有宽列绘制都算作此路径。底层私有 scratch bitmap 的 Lock/Unlock 亦不能直接当作 compositor 阻塞证据。
- 超大单元格的缓存键包含可见裁剪偏移，超过 4096 像素或 1,048,576 像素面积时，逐像素滚动可能持续生成新 tile；样本的约 1768/1049 DIP 宽列与普通行高在本次环境未触发这一阈值，不据此认定样本根因。
- 横向滑块由 WPF `Thumb` 捕获鼠标；UI 线程若迟迟不能处理 MouseUp，释放捕获也会延迟。但 Windows 文档允许用户在捕获期间点击其他窗口，不能仅凭捕获机制解释 Explorer 也不响应。

增加进程级 `--software-rendering` 对照入口，须在创建窗口前设置 `RenderOptions.ProcessRenderMode = SoftwareOnly`；默认启动不变。`Start-QuickExcel-Software.cmd` 使用 Release 配置启动该模式。`test-m5-grid.ps1 -SoftwareRendering` 会核对新的 `rendering.json`，记录实际配置、能力等级、运行时与 OS；能力等级不代表该进程实际使用 GPU。

本轮 Desktop Release 构建成功，仍有 NU1900 数据源警告。当前 Debug 实例占用 DLL，Debug 重建因 MSB3027/MSB3021 文件锁失败，未关闭用户实例。默认和软件渲染下 GridIntegration 均通过，证据分别在 `artifacts/validation/m5-grid/20261030-020728-0bc7ee45/GridIntegration/` 与 `20261030-020634-654c5631/GridIntegration/`。软件模式首轮暴露 smoke 中 2×2 网格误用 B2:C3 合并范围的测试夹具问题，已改用独立小表；失败证据 `20261030-020443-99c44748` 保留。

软件渲染的原文件六场景压力检查位于 `20261030-020805-abdc10d4/HorizontalScrollBenchmark/`，全部到达终点且原文件 SHA-256 未变。两张表宽列场景内容提交间隔 p95 为 26.01/26.27 ms，首次全宽场景仍出现约 490 ms 间隔；不能据此声称软件渲染解决停顿。此次压力检查直接投递 WPF ScrollEvent，未覆盖原生 WM_MOUSEMOVE/WM_LBUTTONUP 与 Thumb 捕获流程。要确定用户描述的数秒停顿，仍需卡住时 UI/渲染线程栈，以及原生输入到达和 Dispatcher 处理时序。未运行完整回归，未关闭物理输入或 M6 性能门禁。

框架依据：[Dispatcher 消息合并](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/WindowsBase/System/Windows/Threading/Dispatcher.cs)、[MediaContext 帧内布局](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/PresentationCore/System/Windows/Media/MediaContext.cs)、[Image.Source 元数据](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/PresentationFramework/System/Windows/Controls/Image.cs)、[BitmapVisualManager 同步通道](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/PresentationCore/System/Windows/Media/Imaging/BitmapVisualManager.cs)、[Thumb 捕获生命周期](https://github.com/dotnet/wpf/blob/v10.0.0/src/Microsoft.DotNet.Wpf/src/PresentationFramework/System/Windows/Controls/Primitives/Thumb.cs)、[SetCapture 跨窗口语义](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setcapture)。

### 横向滑块持续拖动的帧节拍修复（2026-10-30）

用户再次指出“直接点击位置和左右箭头正常，只有鼠标拖动卡顿，宽列尤其明显”。在已有事件合并和矢量缓存基础上，补测 `笼卡demo.xlsx` 的 I 列连续移动：原来的 16 ms Background 定时器与 WPF 合成帧不同步，定时器触发后还要等绘制提交，使内容更新频率明显低于输入频率。此前只统计事件处理/排队的测试无法发现这种画面停顿。

`ThumbTrack` 现在只覆盖待处理位置，在 `CompositionTarget.Rendering` 中每帧至多提交一次最新位置，并在同一帧提交网格内容；相同 `RenderingTime` 的重复回调不重复工作。50 ms Background 定时器只作为窗口被遮挡、合成停止时的后备。完成或取消会同时解除 Rendering 订阅并停表，松手、点击、程序导航、重新绑定和卸载仍保留原有立即提交/取消语义。

`HorizontalScrollBenchmark -WorkbookPath` 新增最宽列场景：独立后台每 8 ms 提交一次输入，四次往返共 241 个请求；在 WPF Rendering 回调观察内容版本变化，保存帧时间、间隔、列位置及全部输入排队样本。计时是 WPF 内容提交节拍，**不是物理显示帧率或鼠标到显示器延迟**，不丢弃冷启动样本。

| 最宽列场景 | 修复前内容更新数 / p95 间隔 | 最终修复后内容更新数 / p95 间隔 | 最终输入排队 p95 |
| --- | ---: | ---: | ---: |
| Sheet1（约 2.1 秒） | 80 / 38.18 ms | 120 / 25.94 ms | 0.78 ms |
| 模板（约 2.1 秒） | 76 / 45.64 ms | 118 / 25.42 ms | 0.95 ms |

原始对照位于 `artifacts/validation/m5-grid/20261030-014419-df2ba4f3/HorizontalScrollBenchmark/`；最终结果位于 `20261030-014855-dff47feb/HorizontalScrollBenchmark/`，中间测量 `20261030-014538-e07e966d/` 亦保留。最终全宽/宽列六个场景均到达请求终点，源文件 SHA-256 未变。首次全宽移动仍记录到约 244 ms 的冷路径停顿（基线约 286 ms），因此这些结果只证明持续拖动节拍改善，不能声称所有冷路径卡顿已消失或稳定 60 FPS。

Debug / Release 构建通过，仅有既有 NU1900 警告。`GridScrollRangeTests|OversizedGridCellTests` 13 项通过；`GridIntegration` 通过，横向专项从 26 增至 31 项，新增同帧内容提交、重复帧回调去重、最新位置覆盖、无延后重播、卸载取消检查，保留冻结/隐藏/合并/编辑器及缓存像素一致性验证。WPF 证据位于 `20261030-014806-140e9173/GridIntegration/`。未运行完整回归。

### 横向拖动回退修正与实际文件验证（2026-10-30）

首轮改动后用户报告打开“笼卡demo.xlsx”拖动更卡，拖到最后可能卡死。检查发现首轮格式缓存对新出现的单元格在 `OnRender` 内逐格执行 `RenderTargetBitmap.Render`，把同步栅格化和原生位图分配放到了 UI 线程；底层网格还重复绘制随后被格式层盖住的文字。原先的小步基准每步等绘制结束才继续，无法测出连续输入积压，不能作为该文件拖动体验的验收依据。

本轮撤掉逐格位图缓存，只缓存冻结的矢量 `DrawingGroup`，按可见范围淘汰；格式层负责的文字不再在底层重复绘制。横向 `ThumbTrack` 仅保存最新位置，由 16 ms、Background 优先级的 DispatcherTimer 合并处理，释放滑块的 EndScroll 立即提交最终位置。尺寸变化、重新绑定、程序导航与卸载会取消待处理位置，避免过期输入覆盖新状态。保留像素滚动及实际宽度终点计算。

新增 `-WorkbookPath` 基准入口，在不写回源文件的条件下遍历工作簿的每个工作表。后台独立以 16 ms / 2 ms 间隔发送各 241 个真实 WPF `ThumbTrack` 路由请求，覆盖四次全宽移动和最终右端位置；保留所有调度排队和处理时间样本。“笼卡demo.xlsx”的两个工作表有明显不等宽列（I 列约 1768 / 1049 DIP），适合补充原先等宽、小幅测试的空缺。测试前后源文件 SHA-256 均为 `21E8714CBB95CA1105B50E699E21246CC364E6B00147C0F1B614BFEB60F518F6`。

| 2 ms 请求场景 | 修正前 Release 排队 p95 | 修正后 Release 排队 p95 | 修正后 Debug 排队 p95 |
| --- | ---: | ---: | ---: |
| Sheet1 | 48.09 ms | 14.46 ms | 29.77 ms |
| 模板 | 62.72 ms | 13.78 ms | 12.52 ms |

Release 修正前、后原始样本分别位于 `artifacts/validation/horizontal-regression-before/`、`artifacts/validation/horizontal-regression-after/`；Debug 为 `artifacts/validation/m5-grid/20261030-012832-f209d721/HorizontalScrollBenchmark/`。所有场景最终位置正确，释放并完成绘制后的输入探针仍能响应。该探针不是“松手到显示完成”的耗时；调度排队也不是物理显示延迟。16 ms 请求下 Release 排队 p95 修正前为 1.47 / 0.69 ms，修正后为 5.43 / 2.28 ms，并非所有指标都降低。此次没有稳定复现永久死锁，故不将这些结果表述为已证明消除所有卡死或达到稳定 60 FPS。

修正后的 Debug 另经 Windows SendInput 验证：打开同一源文件，两张表分别将底部滑块拖到最右端、反向拖回，之后切换工作表仍正常；源文件哈希未变。此项验证不是人工物理鼠标验收。Debug / Release 构建成功，仍有既有 NU1900 数据源警告。定向 WPF 检查通过：GridIntegration（横向专项从 21 增至 **26** 项，覆盖突发输入合并、释放提交、过期位置取消）、Workspace **151** 项、Interaction **92** 项、DataOperations **81** 项。对应目录为 `20261030-012748-1e331d12`、`20261030-012759-4bae67cb`、`20261030-012810-864cb3d3`、`20261030-013157-2bfe6b2d`。未运行完整回归。

```powershell
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode HorizontalScrollBenchmark -WorkbookPath 'D:\Downloads\笼卡demo.xlsx'
```

### 首轮底部横向滑块改动（历史结果，已由上节修正）

原先底部滑块把 `Value` 四舍五入成列索引，已有像素滚动只用于拖选自动滚动。因此小幅拖动时内容停住，越过半列后突然跳一列；`ScrollChanged` 还在实际滚动之外重复通知视口，即使没有移动也会触发格式重绘。`FastGridArtwork` 每帧重新创建可见格式单元格的 `FormattedText`，重复格式化、排版及提交字形。

首轮将 retained 模式的横向滑块改为以设备像素计量，保留首列内偏移，按实际列宽计算终点；纵向滚动保留横向偏移，窗口变宽会校正右端位置。内容变化只发一次视口通知，未移动的 EndScroll 不重绘。旧 bitmap 模式继续按列滚动。首轮格式层按可见地址、尺寸和 DPI 缓存不可变绘制结果，其中普通单元格逐格栅格化、超大单元格保留矢量绘制；逐格栅格化已在上述回退修正中移除。快照刷新清缓存，离开视口的内容被淘汰，选区、编辑框和图片仍按各自分区绘制。

新增独立 `HorizontalScrollBenchmark`，使用 100×100 已填充单元格、100,000×256 逻辑网格和 1400×850 DIP 窗口，以真实 WPF `ThumbTrack` 处理器测试每次 1/8 列的移动。5 次预热、60 个原始计时样本；模型读取和分配量包含预热的 65 次操作。相同循环原先只有 6 个不同显示位置，修复后有 40 个。计时终点为 Dispatcher/内容与选区重绘完成，不是物理鼠标到显示器延迟。

| 指标 | 首轮前 | 首轮结束（非当前版本） |
| --- | ---: | ---: |
| 普通内容重绘 p95 | 8.78 ms | 14.20 ms |
| 带格式内容重绘 p50 | 24.37 ms | 6.13 ms |
| 带格式内容重绘 p95 | 47.50 ms | 26.35 ms |
| 带格式场景累计托管分配 | 222.62 MB | 57.57 MB |
| 带格式场景模型读取 | 20,025 | 1,050 |

普通内容现在每个像素步骤都会移动和绘制，不能与原先多数步骤停住的耗时直接解释为同等工作量。带格式 p95 仍超过预先声明的 16.67 ms 预算，不声称稳定 60 FPS 或关闭 M6。中间轮 p95 为 18.38 ms，也保留原始数据，不以较好的一轮代替最终数据。隐藏格式层的对照只移除合成；WPF 仍可能调用 `OnRender`，不能解释为完全移除格式计算。

基线证据：`artifacts/validation/m5-grid/20261030-010506-2952a5b7/HorizontalScrollBenchmark/`；最终：`20261030-011341-cad19f53/HorizontalScrollBenchmark/`。中间轮为 `010711-b0c445e5`（仅复用文字布局）及 `010846-6aad381e`（加入栅格缓存），均保留 JSON 原始样本。

Desktop Release 构建成功，只有既有 NU1900 数据源警告。`GridScrollRangeTests|OversizedGridCellTests` **13 项通过**；Interaction **92 项**、Workspace **151 项**通过，目录分别为 `20261030-011157-e15eb26c`、`20261030-011200-780a0f5c`。最终 GridIntegration 位于 `20261030-011328-77eb1739`，新增 **21 项横向专项断言**覆盖单像素位移、零重复读取/排版、通知次数、冻结/隐藏/合并、编辑裁剪、缓存有界、快照失效、超大单元格、两种模式切换和右侧终点；缓存与清空缓存后的像素一致。未运行完整回归，未声称物理鼠标、跨屏 DPI 或完整 OS IME 验收。

```powershell
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode HorizontalScrollBenchmark
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode GridIntegration
```

### 双击自动调整列宽／行高（2026-10-30）

双击列标题右边界按显示内容自动调整该列宽度；双击行标题下边界按当前列宽、字体和换行设置自动调整该行高度。测量使用与网格相同的普通字形／样式文字布局及数字日期格式，包括视口外的已存储单元格；跳过手动隐藏、筛选隐藏及合并单元格，不计图片。列宽以最长显示行测量，空轴恢复默认尺寸，上限沿用 10,000 DIP。测量按稀疏内容遍历，使用有界文字尺寸缓存；合并范围按轴排序后顺序跳过。

先提交原地址草稿，再读取新快照测量；纯尺寸变化只提交一个可撤销命令，不启动重算、不改变选区／筛选／冻结，也不授予脚本权限。重复双击相同结果不增加历史；输入法组合和忙碌状态沿用现有门禁。普通单击拖动保留。

Desktop Release 构建成功，只有既有 NU1900 网络警告。本轮 Workspace **151 项断言通过**（新增 24 项自动尺寸检查），证据在 `artifacts/validation/m5-grid/20261030-004527-f788623f/Workspace/`，含 `auto-fit.xlsq` 和 `auto-fit.png`；Interaction **92 项通过**，证据在 `20261030-004546-e1a17d40/Interaction/`。另以 Windows SendInput 在新建空白工作簿验证列／行双击以及之后继续拖动，未操作已有用户文件；该检查不等同于人工物理输入验收。未运行完整回归或性能基准。

### 手动列宽／行高调整（2026-10-30）

行列标题边缘现在显示横向／纵向调整光标；主窗口仅在调整边缘或拖动期间设置 `ForceCursor`，避免子网格的普通箭头覆盖提示。拖动实时预览，松手按最终指针位置提交一次尺寸命令；Escape、窗口失活、捕获丢失或捕获失败均取消预览并释放光标，不新增尺寸历史。既有原格草稿提交、输入法门禁、撤销重做和保存重开保持。

标题命中改为网格独立的视图几何接口，按设备比例及冻结／滚动分区计算可见边缘，不再依赖另一轴的单元格是否可见。所有行隐藏后仍可调整列宽，所有列隐藏后仍可调整行高；很矮的行取最近边缘，隐藏项、视口裁切边、滚动条和左上角不产生错误目标。文档写入仍走 `EditorController` / `DocumentSession`。

Desktop Release 构建通过；`M5DimensionTests|GridScrollRangeTests|OversizedGridCellTests` 定向 **18 项通过**；Workspace **127 项 WPF 断言通过**，覆盖光标路由、取消、尺寸保存、冻结滚动、隐藏轴及窄行边界。最终证据在 `artifacts/validation/m5-grid/20261030-003555-e5cbf975/Workspace/`，包括 `resize-headers.png`。隐藏 smoke 窗口的捕获丢失使用合成 WPF 路由；另用 Windows SendInput 在空白工作簿实际拖动两轴并核对显示结果，未操作已有用户文件。这不是人工物理鼠标或完整 OS IME 验收。未运行完整回归或性能基准，构建仍有 NU1900 漏洞数据源不可达警告。

### 筛选后滚动范围修复（2026-10-30）

纵向滚动范围此前按默认行高估算，忽略导入或自定义的实际行高；筛选后剩余记录较少、实际高度仍超出视口时，滚动条可能被压到零，无法到达下方数据。现在行列共用可见尺寸索引计算滚动终点，排除隐藏与冻结项，并处理整页恰好容纳、空结果和超大单元格。范围缩小时先校正视口索引，再更新滑块位置和可见数量，避免沿用筛选前的越界位置。

新增 `GridScrollRangeTests` 验证尺寸边界；`GridIntegration` 的筛选专项通过真实筛选处理函数及 WPF 滚动事件检查 11 项行为，包括较高行的底部可达性、滚轮、区域外空行、冻结表头、空结果、重新筛选和清除条件。旧实现复现失败证据在 `artifacts/validation/m5-grid/20261030-001501-5f2fc1f4/GridIntegration/smoke-error.txt`；修复后的结果和底部截图在 `20261030-001717-8a01efac/GridIntegration/`。

Desktop Release 构建通过；`GridScrollRangeTests|OversizedGridCellTests|DataOperationsTests` 共 32 项通过；`GridIntegration` 通过；`DataOperations` 的 81 项断言通过，证据在 `20261030-001745-a43df4df/DataOperations/`。所有托管输出均在 `artifacts/build`，未运行完整回归或性能基准。构建仍有 NuGet 漏洞数据源不可达的 NU1900 警告。这些检查验证 WPF 处理路径，不代表物理鼠标输入验收。

### 选框重叠修复（2026-10-29）

普通多格选区内的小框有两个来源：选区层重复绘制当前格边框，以及透明原生 TextBox 的 Windows 主题焦点边框。现在当前格已在选区内时只绘制选区外框；待输入的文本框保持键盘焦点，但整体透明，进入编辑时恢复显示。当前地址和公式栏保持原语义。

先用 2×2 数据 `7/8/11/12` 复现，旧实现的选框断言失败证据在 `artifacts/validation/m5-grid/20261029-081852-ffbf06af/GridIntegration/`。最终 Desktop Release 构建通过，`GridIntegration` 的 11 项选框/编辑器检查和 41 项合并选区断言通过；整网格像素对比确认待输入文本框没有焦点边框残留，截图在 `20261029-082111-f2489372/GridIntegration/selection-outline.png`。`Interaction` 92 项断言通过，证据在 `20261029-082125-1ed9c382/Interaction/`。未运行完整回归；构建有已知 NU1900 数据源警告。验证为 WPF 绘制及处理路径检查，不代表物理输入验收。

### 合并单元格选区修复（2026-10-29）

合并内容层与选区层此前使用不同的范围：选中 F1:H7 时仍绘制 F1 的小框。现在网格接收快照中的合并范围，点击、拖选、程序选区和 Ctrl 增删均覆盖整个合并区域；当前格边框也使用合并范围。编辑地址归属左上角，编辑框覆盖合并区域，方向键、Tab/Enter 从区域边缘移出。单个合并区域粘贴单值只写左上角，并保持一次撤销。

针对本次修复，Desktop Release 构建通过，`MergedSelectionTests|M5InteractionTests` 共 24 项通过；未运行完整回归。`GridIntegration` 新增 41 项合并选区断言，涵盖 F1:H7、内部命中路径、Ctrl 增删、拖选、编辑及粘贴归属、导航、冻结滚动裁剪、隐藏行、取消合并与撤销。证据与截图在 `artifacts/validation/m5-grid/20261029-080318-75280ca9/GridIntegration/`。`Interaction` 92 项断言通过，证据在 `20261029-080325-c2e99551/Interaction/`。点击检查调用真实命中及选区处理路径，不代表物理鼠标输入验证。构建仍有 NuGet 漏洞数据源不可达的 NU1900 警告。

### 主网格迁移验证

所有命令从仓库根目录执行，构建输出在 `artifacts/build`。没有运行完整回归套件。

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter "FullyQualifiedName~M5InteractionTests|FullyQualifiedName~M5DimensionTests|FullyQualifiedName~M5WorkspaceTests|FullyQualifiedName~DataOperationsTests|FullyQualifiedName~PresentationCalculationTests" -v minimal
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Interaction
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Workspace
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Ribbon
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Product
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode DataOperations
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode GridIntegration
```

Release 编译通过，NuGet 漏洞数据源不可达产生 NU1900 警告。78 项相关 Core 测试通过。专项 WPF 断言包括 Interaction 92、Workspace 108、Ribbon 21、Product 41、DataOperations 81；新增 GridIntegration 独立覆盖双向编辑归属、无效草稿地址、编辑中保存、日期数值精度和真实 WPF 关闭生命周期，避免为这些检查运行完整脚本/存储回归。Product 的 AppContainer 部分在受限执行器失败，在正常 Windows 用户环境通过，未弱化隔离。

原始功能证据位于 `artifacts/validation/m5-grid/` 的各次独立运行目录；`grid-core-tests.log` 和 `grid-solution-build.log` 位于 artifacts 根目录。

| 模式 | 通过的证据目录（相对 artifacts/validation/m5-grid） |
| --- | --- |
| Interaction | `20261029-074600-c36ae3c1/Interaction` |
| Workspace | `20261029-074602-5fd07272/Workspace` |
| Ribbon | `20261029-074407-ec13959b/Ribbon` |
| Product | `20261029-074411-33ccd7f0/Product` |
| DataOperations | `20261029-074415-637f91c7/DataOperations` |
| GridIntegration | `20261029-074607-d16805b7/GridIntegration` |

性能单独执行，保留原始样本：

```powershell
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode FastGridBenchmark
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode FastGridDragBenchmark
```

`20261029-073638-5290aabe/FastGridBenchmark` 达到预设预算：导航 p95 30.23 ms，25.6M 格选区 p95 1.28 ms，100k 行筛选 p95 82.48 ms。绑定只有一个样本，不视作分位数。`20261029-073839-2eb188e6/FastGridDragBenchmark` 的绘制预算通过，但像素自动滚动节拍 p95 为 33.66 ms，高于 33.34 ms 门槛，整体预算未通过；不能把结果文件中的“PASS: benchmark completed”解释为所有性能门槛通过。

移除无格式工作簿的多余样式遍历、复用一次可见列列表后，`20261029-074054-0f6b3814/FastGridDragBenchmark` 整体预算通过：视口拖选绘制 p95 2.50 ms、0 模型读取；像素滚动节拍 p95 32.28 ms。两次原始结果均保留，调度节拍接近门槛且受本机环境影响，不能据单次通过关闭 M6。

真实 OS IME 候选窗口、跨屏 DPI、完整 UI Automation 单元格导航和 M5/M6 性能/安全门禁仍需独立验收。当前路由事件、模型选区、像素快照和 Dispatcher 时序检查不代表物理输入到显示延迟。默认切换是本次用户明确要求的集成变更，不据此关闭上述质量门禁。
