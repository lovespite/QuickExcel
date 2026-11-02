# E4 报表、打印与交付能力

2026-11-02：XC-0401–0405 完成下列声明范围。实现基于 `8ec23d7`，本轮变更在工作区，尚未提交；未生成发布包，未关闭 E5、M4/M6 或真实物理打印机/跨屏输入门禁。所有代码和证据均在 QuickExcel 仓库，ExcelIO 源码未修改。

## 支持矩阵

| 项目 | R / D / E / W | 本轮交付 |
|---|---|---|
| XC-0401 | 读、预览、设置、输出 | 报表页的页面设置、分页预览、系统打印和 PDF；复用已有打印元数据 |
| XC-0402 | 读、显示、编辑、写回 | 三类浮动图片锚点、EMU 偏移、移动/尺寸/删除、跟随单元格规则 |
| XC-0403 | 读、显示、编辑、写回 | 簇状柱形、基础折线、饼图；标题、系列、数据范围、图例及基础数值轴 |
| XC-0404 | 读、约束/显示、编辑、写回 | 工作表保护、锁定/隐藏公式、允许操作、传统批注增删改 |
| XC-0405 | 分级预算、真实测量 | 200,000 行 × 2,048 列坐标/网格、8,192 样式、512 MiB 解析计费 |

新增编辑仍通过 DocumentSession 一次提交、撤销和重做。未支持的依赖和附件不由本轮放开；现有只读原包、明确另存受支持副本、保存冲突与重开校验继续生效。

## XC-0401 打印与 PDF

Core 的 PrintingLayout 在 96 DIP 坐标中有界分页，WPF WorkbookPaginator 同时驱动预览、系统 PrintDialog 和 PDF。支持 A4/A3/A5/Letter/Legal、已有自定义纸张、横纵向、页边距、10–400% 缩放、适应宽/高页数、打印区域、多区域、重复行列标题、完整跨度手动分页、两种分页顺序、居中、网格线和行列标题。适页按整百分比缩小，不把较小报表扩大到 100% 以上。手动隐藏和筛选隐藏均排除，重复标题只在需要的后续页补入。

普通格式、数字显示、合并、富文本、旋转/缩进/缩小填充、基础边框、图片和基础图表进入同一分页器。黑白模式将文字/单元格样式转为黑白，图片/图表转为灰度。页眉页脚支持普通文本、`&&`、`&L/&C/&R`、`&P/&N`、`&A/&F`、`&D/&T`，包括首页/奇偶页和首个页码；保留并应用随文档缩放及页边距对齐开关。

打印区域最多 2,000,000 格、任务最多 2,000 页，重复标题最多 256 行/64 列。空可见区域、单行列连同标题超过一页、局部跨度分页、批注打印、未支持的页眉字体/图片/路径指令会明确拒绝。含待重算公式或未支持绘图诊断时不生成可能误导的打印结果。未知纸张不能静默改为 A4。图片完整打印解码上限为 1,600 万像素。

PDF 使用已安装的 Windows **Microsoft Print to PDF** 驱动，经 XPS 和 XpsPrint 写到唯一临时文件，完成并检查 PDF 文件头后替换目标。XPS 预算 64 MiB、打印完成等待上限 60 秒；缺少驱动或失败会报告。这里验证的是本机驱动，不是内置、跨平台 PDF 引擎，也没有向物理打印机发送纸张。系统打印由用户选队列；实际设备的不可打印边距和驱动行为仍需设备验收。

三份报表分别输出 1、2、2 页。手动分页表在 Excel 和 QuickExcel PDF 中的每页 Record 集合完全相同：第二页重复首行后从工作表第 21 行开始。96/144/192 DPI 的页面尺寸、页数和裁剪一致；缩回 96 DPI 后通道差大于 32 的像素占比最高 0.688%，主要是字体/边缘抗锯齿，阈值为 5%。已目视检查普通表、合并、双色图片、跨页图表和 PDF 页面。自动图表报表双方均两页，但 Excel 的字体、自动坐标刻度、图表内部排版和驱动度量不与 WPF 逐像素一致；不把此结果宣称为完整 Excel 印刷复刻或真实显示器 DPI 切换验收。

## XC-0402 图片

DrawingAnchor 保存整数 EMU（96 DPI 下每像素 9,525 EMU），支持 twoCell、oneCell、absolute。Excel 常见的 twoCell + editAs 读入后归一为相应行为：随单元格移动并缩放、只移动、固定位置和尺寸。报表页可插入 PNG/JPG/GIF/BMP/TIFF，输入带小数的工作表位置和宽高，改变跟随规则或删除。

行列插入、删除和尺寸变化按锚点行为变换，保留偏移；完全被删掉的 twoCell 图片移除，删除 oneCell 起点的情形明确拒绝。带图表、分组、批注/链接等现有限制仍适用；不为图片放松未知附件的结构门禁。裁剪、旋转、单元格内图片和复杂绘图继续保留/报告，不能宣称其编辑保真。单次 UI 插入最多 16 MiB，每表图片最多 1,024 张、64 MiB；还受整包预算约束。

本轮补齐默认列宽保存状态，区分 Excel 省略的默认列宽与显式 defaultColWidth。二者不能互换，否则 Excel 会改变列宽并使 twoCell 图片漂移。插删时同时保留新增 Data 元数据。字体最大数字宽度和复杂列宽度量仍沿用既有布局边界，未宣称支持任意 Normal 字体下的像素等价。

真实 Excel 对照的三种图片行为在写回及 Excel 再保存重开后保持；位置和尺寸与输出模型差不超过 1 point。独立 Pictures 夹具验证插入→删除恢复原锚点，再插入后写回。组合夹具含批注 VML，所以结构测试在独立图片工作簿执行，组合文件通过受支持的图片位置编辑命令应用结果，没有绕过产品的批注结构保护。

## XC-0403 图表

基础子集为纵向簇状柱形、无标记/无平滑基础折线、单系列饼图。数据直接引用本工作簿单表的单行/单列向量，可跨表；分类和值等长，每系列最多 4,096 点、每图最多 16 系列，每表最多 128 图及 8 MiB 定义。UI 以首行系列标题、首列分类的矩形创建或更改数据范围，饼图需要两列；可修改标题、类型、图例位置或删除。

读写基础 min/max、数值格式、系列颜色、饼图点颜色、四侧/隐藏图例和 plotVisOnly。网格与打印共用向量绘制，从当前快照取数据，分类使用显示格式；隐藏/筛选数据按设置排除。序列名称、分类、数值缓存均在保存时由源引用重建，缓存包含完整源序号，不能把隐藏项压缩后写成错位缓存。源公式未重算时网格显示原因，保存/打印拒绝旧缓存。工作表改名和复制同步更新图表引用；受影响的行列插删/搬移保守拒绝。

堆积、水平、组合、3D、趋势线、误差线、数据标签、非默认标记/平滑、日期/对数/多组轴、自定义布局/富样式、外部数据、名称/动态引用及未知扩展不属于此子集，进入原包保留/报告路径；不显示旧图形。XML 白名单只接受已理解结构，已知无依赖的 Office 身份/NA 显示扩展单独验证。基础图表的字体、自动刻度、图例裁剪和版式采用本应用绘制，不复刻 Excel 所有图表主题。已有 Excel 的 Office 2017 dispNaAsBlank 导致 SDK 校验 6 条发现：工具保留原始记录，只允许位置和消息均与源包完全相同的发现，不忽略新增结构错误。

## XC-0404 保护与批注

样式记录 locked/hidden 并保持 XF 继承。DocumentSession 统一检查保护，覆盖锁定内容写入、批量混合写、格式、行列尺寸、插删、未锁定区域排序、现有筛选条件以及图片/图表/批注；失败不改变快照和历史。选择限制和公式栏隐藏在桌面投影；内部复制受保护的隐藏公式只提供值。允许格式操作不能修改锁定/隐藏属性。行列删除还检查空白默认锁定和显式锁定覆盖；部分尚不支持的结构组合即使保护标志允许，也继续明确拒绝。

UI 可设置单元格锁定/隐藏、启用/取消无密码工作表保护，调整展示的允许操作，未展示的原标志仍保留。导入的传统密码或现代哈希属性读写保留，不能在本应用中取消或更改这种保护，需在 Excel 处理。工作表保护只是操作约束，不是文件加密或保密机制；工作簿结构保护、允许编辑区域、文件密码解密未纳入本轮。

传统批注可创建、修改作者/文本、删除和撤销。替换含 run 格式的批注须勾选明确转换为普通文本；不改变位置的文字编辑保留原 VML 字节，新增/删除仅协调经过验证的 note-only VML 形状，保留其他批注形状。线程评论和其他 VML 对象仍不支持。图片数据、批注/VML、样式、打印设置和图表定义计入历史成本。

## XC-0405 容量分级

本机 Windows x64、.NET 10.0.12、i9-13980HX、约 31.6 GiB 物理内存。基准与测试/WPF 分开执行；每个稠密级别 10 列连续数值，三个原始样本，覆盖构造、保存含重开校验、重新打开、20 次单格编辑及撤销、1 ms 后请求取消。

| 固定负载 | 打开 ms | 保存 ms | 20 次编辑 ms | 进程累计峰值工作集 MiB |
|---|---:|---:|---:|---:|
| 25,000 存储格 | 121–171 | 285–437 | 143–371 | 113–126 |
| 100,000 存储格 | 274–320 | 510–575 | 619–675 | 243–259 |
| 250,000 存储格 | 811–889 | 1,347–1,386 | 2,186–2,411 | 474–579 |

所有 20 次编辑历史计费均为 35,840 bytes；所有取消观察时间为 0.26–2.12 ms。峰值是同进程高水位，live delta 和累计分配另存原始 JSON，不能混称稳态 RSS 或跨机器承诺。最初 256 MiB 计费预算在 250,000 格保存重开时拒绝；之后将解析计费独立提升到 512 MiB，同时提高写出单格保守预检成本，保留 64 MiB 文件/条目及 256 MiB 展开上限。

4,096 / 8,192 个不同字体颜色样式各测三次。8,192 样式打开 111–152 ms、构造加保存 190–213 ms，样本文件 97,845 bytes；本轮累计分配约 335 MiB，不是保留内存。边界坐标 (200000, 2048) 仅存一格，文件 2,219 bytes，保存重开 21.9 ms，并由真实 WPF 导航到该格。

拒绝边界同时生效：

- 每表坐标 200,000 行 × 2,048 列，网格可到达全部列；全簿最多 64 表、1,000,000 个存储格、8,192 样式。命名样式等独立既有预算不随之提升。
- 文件、单条目和输出 64 MiB；累计展开 256 MiB；XML/JSON 和对象解析计费 512 MiB。512 MiB 是计费预算，不是进程内存上限。文本文字、对象和元数据会提前耗尽预算；百万格是硬上限，不是已通过的稠密负载等级。
- XML/JSON 深度、ZIP 中央目录、图片和脚本限制保持。单次粘贴 20,000 格、复制 100,000 格、填充/格式等原有操作预算保持；200,000 × 2,048 不代表可以一次填满此矩形。
- 大图、图表、过滤、历史、打印分别计费；超限抛出明确错误，不截断内容或提交部分结果。这里没有测满百万格、复杂公式/脚本、极端多图、低内存设备、网络盘或 M6 并发压力。

XLSQ 继续 formatVersion/minReaderVersion 2，按内容增加 `drawing-anchors`、`worksheet-protection`、`worksheet-charts`、`worksheet-default-width`、`extended-capacity` 必需能力；旧读取器必须按已有未知能力门禁拒绝，不能悄悄忽略保护、宽度或对象语义。

## 验证及复现

- Release solution 和最后 Desktop 构建成功。NuGet 漏洞源连接失败产生 NU1900；使用已还原包构建，未关闭警告或安全检查。
- 289 项相关 Core/存储/结构/剪贴板/交互/过滤/样式回归通过；收尾后再跑受改动影响的 117 项通过，其中 E4 专项 23 项。没有运行 worker/原生/全量回归。
- Excel 16.0 build 20527、语言 2052：四表夹具、81 项普通打开/保存/重开检查通过，没有进入 RepairMode；图表源改为 123.5，保护表未锁定值改为 77，批注改/增/删可回读。输入 SHA-256 不变。
- Excel 原始文件仍有既有 phonetic、uid、workbook 扩展等诊断。原生夹具的编辑输出明确使用受支持副本；并未把这种文件标为完全无损。另有 WPF 新建报表的普通保存/重开成功。
- 最终 WPF **23 项**真实事件/对话框检查通过，含图片尺寸、图表插入、批注、保护、页面设置、保存重开、PDF 及新坐标边界。自动触发事件不等于物理键鼠/IME 验收。

最终证据（均为 ignored artifacts）：

- `artifacts/e4/tests/e4-related-final.trx`、`e4-final-focused.trx`。
- `artifacts/e4/excel-final-10/`：原始和编辑 XLSX/XLSQ、独立图片结构夹具、源/输出 SDK 审计、Excel 再保存文件、两组 Excel PDF、81 项记录及哈希。
- `artifacts/validation/m5-grid/20261102-064631-fe2b66b5/Reporting/`：三份 PDF、每份首张的 96/144/192 DPI PNG、跨页 PNG、布局 JSON、网格/容量截图和 `visual-validation.json`；父目录 `run.json` 记录二进制哈希与环境。
- `artifacts/e4/capacity-before/`、`capacity-after/`、`capacity-final/`：保留初始、调整后及最终样本；最终为九份稠密、六份样式样本和一份稀疏边界。
- `excel-final-01`–`06`、`wpf-01`–`02` 的失败留存，包括探针参数/隔离夹具问题、默认列宽漂移、PDF ticket 和旧 UI 列限制；未删除失败证据或把它们当作通过。

从仓库根目录复现，输出须使用新的 artifacts 目录：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter 'FullyQualifiedName~E4ReportingTests|FullyQualifiedName~WorksheetPrintingTests|FullyQualifiedName~TypedStorageTests' -v minimal
.\scripts\test-excel-e4.ps1 -OutputDirectory .\artifacts\e4-new-excel
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Reporting -WorkbookPath .\artifacts\e4-new-excel\excel-source.xlsx
# 独立运行基准；不要与测试/WPF 并发。
dotnet run --project .\tools\QuickExcel.CompatibilityProbe\QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path .\artifacts\build -- e4-capacity .\artifacts\e4-new-capacity
# Python 环境需要 Pillow、pypdf；第一个参数替换为上一步输出的 Reporting 目录。
python .\scripts\verify-e4-visuals.py <Reporting目录> .\artifacts\e4-new-excel
```

打印入口及保护语义参考 [Microsoft XpsPrint](https://learn.microsoft.com/en-us/windows/win32/api/xpsprint/nf-xpsprint-startxpsprintjob)、[Excel 工作表保护](https://support.microsoft.com/en-us/excel/protect-a-worksheet)；列宽与字体度量边界参考 [OOXML Column](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column)。
