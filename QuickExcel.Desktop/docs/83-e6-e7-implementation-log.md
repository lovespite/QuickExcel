# E6/E7 日常功能实施记录

## XC-0601a — CSV/TSV 导出 / CSV and TSV export

2026-11-05，产品版本 1.17.0。完成 UTF-8 CSV/TSV 导出的独立子项，导入预览属于后续 XC-0601b。

文件菜单提供当前工作表（从 A1 到已用范围右下角）或单一矩形选区导出；逗号/制表符、可选 BOM、原始值/显示文本均可选。包含隐藏行列，格式、多表和对象不导出。原始数字使用不依赖区域设置的 round-trip 表示，日期为序列数；显示文本复用单元格格式器。文本与公式结果不猜测类型，Excel 的自动打开类型推断可能改变长编号，应使用文本导入并明确列类型。

Export is available from the File menu for the current sheet or rectangular selection. UTF-8, optional BOM, comma/tab delimiters and raw/display values are explicit options. Hidden rows and columns are included. Formatting, other sheets and objects are omitted. Formula caches must be current; export preserves document history and save identity.

字段引用遵循成对双引号转义，保留中文、嵌入 CR/LF 和尾部空字段。最多 1,000,000 字段及 32 MiB 字符；逐行取消，目标指纹检查及同目录原子替换。取消、失败或过期公式不覆盖目标文件。多区域导出需先选择单一矩形。

验证：`TextExchangeTests|LocalizationCatalogTests` 共 9 项通过；Avalonia TextExchange 模式 6 项通过（菜单、中文引号换行、BOM、保存身份、取消选项和取消文件选择）。证据：`artifacts/avalonia/workspace/20261105-065800-90f2f999`。未声称本子项已完成真实 Excel 双向文件验收，随后导入子项补充独立样例。

发布：`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 通过；活动包 `artifacts/avalonia/releases/20261105-065818-aca1fe6b/application`，生产自检见同级 `evidence`。已核对 `active-windows.json` 指向此包。源码基线 `ca7d1e1` 加本提交；发布包含原工作区未提交的既有改动，未将这些改动混入本功能提交。
