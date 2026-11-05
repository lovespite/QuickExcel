# E6/E7 日常功能实施记录

## XC-0601a — CSV/TSV 导出 / CSV and TSV export

2026-11-05，产品版本 1.17.0。完成 UTF-8 CSV/TSV 导出的独立子项，导入预览属于后续 XC-0601b。

文件菜单提供当前工作表（从 A1 到已用范围右下角）或单一矩形选区导出；逗号/制表符、可选 BOM、原始值/显示文本均可选。包含隐藏行列，格式、多表和对象不导出。原始数字使用不依赖区域设置的 round-trip 表示，日期为序列数；显示文本复用单元格格式器。文本与公式结果不猜测类型，Excel 的自动打开类型推断可能改变长编号，应使用文本导入并明确列类型。

Export is available from the File menu for the current sheet or rectangular selection. UTF-8, optional BOM, comma/tab delimiters and raw/display values are explicit options. Hidden rows and columns are included. Formatting, other sheets and objects are omitted. Formula caches must be current; export preserves document history and save identity.

字段引用遵循成对双引号转义，保留中文、嵌入 CR/LF 和尾部空字段。最多 1,000,000 字段及 32 MiB 字符；逐行取消，目标指纹检查及同目录原子替换。取消、失败或过期公式不覆盖目标文件。多区域导出需先选择单一矩形。

验证：`TextExchangeTests|LocalizationCatalogTests` 共 9 项通过；Avalonia TextExchange 模式 6 项通过（菜单、中文引号换行、BOM、保存身份、取消选项和取消文件选择）。证据：`artifacts/avalonia/workspace/20261105-065800-90f2f999`。未声称本子项已完成真实 Excel 双向文件验收，随后导入子项补充独立样例。

发布：`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 通过；活动包 `artifacts/avalonia/releases/20261105-065818-aca1fe6b/application`，生产自检见同级 `evidence`。已核对 `active-windows.json` 指向此包。源码基线 `ca7d1e1` 加本提交；发布包含原工作区未提交的既有改动，未将这些改动混入本功能提交。

## XC-0601b — 文本导入预览 / Text import preview

2026-11-05，版本 1.18.0。文件菜单新增“导入 CSV / TSV…”。支持 UTF-8、UTF-16 LE/BE、GB18030 显式编码，以及逗号、制表符、分号、竖线。先预览原文，再逐列选择文本/数字/年月日/月日年/日月年/跳过；支持首行文本标题、区域数字分隔符和 1900/1904 日期系统。转换复用分列规则，拒绝超过 15 位有效数字及无法解释的日期。转换后再次预览，确认后创建新工作簿并要求另存；取消和过期预览保留当前文档。常规“打开 CSV”入口仍保留原来的 UTF-8 文本快速打开行为。

The File menu now offers an explicit import wizard. Choose encoding and delimiter, review the source, assign column types, then review converted values. Text is the default. Cancellation, conversion failures and stale previews preserve the current workbook. Import creates a separate workbook with scripts disabled; saving uses Save As.

容量：输入 64 MiB、解码 32 MiB 字符、最多 1,000,000 字段，最终矩形也受 1,000,000 格限制。保留实际尾部空字段；跳过列保留原列位置。预览展示前 10 行，转换后预览最多前 20 列，所有记录仍经过类型验证。取消在读取后及逐行转换检查；未承诺慢磁盘读取中的实时中断。

验证：19 项 TextImport/TextExchange/Localization 定向测试通过；Avalonia TextExchange 11 项通过，证据 `artifacts/avalonia/workspace/20261105-070306-e411b9b8`。真实 Excel 16.0 build 20527 直接打开 QuickExcel CSV 保留中文、引号和嵌入换行；Excel 导出的 UTF-8 CSV 由 QuickExcel 独立导入并保持编号、等号原文及换行，夹具与 SHA-256 位于 `Fixtures/text-exchange-excel.json`，证据 `artifacts/e6/text-exchange-excel-04`。Excel COM 的旧 OpenText 路径在本机将嵌入 LF 截断，常规 Workbooks.Open 验证通过；两者不能混为同一入口。

发布与自检通过：`artifacts/avalonia/releases/20261105-070355-e1bb7007/application`，已验证活动指针。无新增第三方依赖，GB18030 使用 .NET 已分发的 CodePages provider。
