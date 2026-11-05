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

## XC-0602a — 引用位置函数 / Reference position functions

2026-11-05，1.19.0。新增 ROW/COLUMN 的无参当前坐标、单格及矩形引用，矩形返回行向/列向数组；ROWS/COLUMNS 接受引用、标量及受支持数组表达式。名称、跨表、结构编辑、数组所有权和标准文件公式路径复用现有实现。静态引用的坐标/尺寸不读取被引用格值，防止自引用被误判为循环。无参函数在绑定时补充当前地址，源公式仍保留；可移植导出保留函数语义。提示文案为中英文，参数数量登记接入现有补全。

ROW and COLUMN return current coordinates or reference coordinates, with bounded array results for ranges. ROWS and COLUMNS return reference or supported array dimensions. Static coordinate queries do not depend on cell values. Formula source, structural reference updates, spill conflict checks and undo remain in the existing calculation and document paths.

范围：全列引用按 QuickExcel 的 200,000 行容量解释；数组仍限 100,000 元素。ROW/COLUMN 目前需要静态矩形引用，不接受计算产生的引用数组或溢出运算符作为引用参数；ROWS/COLUMNS 可读取溢出数组尺寸。

验证：36 项引用函数/动态数组/本地化测试通过，补充结构编辑和撤销后 19 项引用专项通过；13 个独立 Excel 对照通过，样例与来源 SHA-256 为 `Fixtures/common-functions-spec.json` / `excel-common-functions.json`。规范参考 [Microsoft ROWS](https://support.microsoft.com/en-us/excel/functions/rows-function) 与 [动态数组说明](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-vs-legacy-cse-array-formulas)。桌面 CommonFunctions 6 项通过，证据 `artifacts/avalonia/workspace/20261105-071129-5d250dad`。

发布与生产自检通过，活动包 `artifacts/avalonia/releases/20261105-071219-533eef9e/application`；已核对启动器活动指针与清单哈希。

## XC-0602b — 逻辑与乘积统计 / Logical and product formulas

2026-11-05，1.20.1（与并行的脚本编辑器 patch 版本合并）。IFS 支持 1–127 个条件/结果对，首个匹配返回结果，未匹配为 #N/A；SWITCH 支持相同类型的精确匹配（文本不区分大小写）、默认值和 #N/A。结果按所选分支求值，不以未选数组扩展结果；数组条件及选中结果按既有单行/单列广播契约组合，尺寸缺失为 #N/A。静态依赖仍保守包含全部源码分支，未承诺脚本副作用与 Excel UDF 调度等价。

SUMPRODUCT 对同形状数组逐项相乘再求和，文本、布尔和空白按零处理，显式算术转换后按数字处理；维度不匹配为 #VALUE!，错误传播，即使另一个乘数为零也不吞掉错误。数组仍有 100,000 元素预算；全列超预算须缩小引用。IFS/SWITCH 的 XLSX `_xlfn` 编码与解码、增量重算、保存重开、撤销和双语参数提示已接入。

IFS and SWITCH select the first matching result and support bounded array conditions and broadcasting. Unselected result arrays do not expand the output. SUMPRODUCT requires equal array shapes, treats nonnumeric entries as zero, and propagates errors. Modern function names round-trip through XLSX.

39 项相关测试、31 个累计独立 Excel 公式样例、8 项 Avalonia CommonFunctions 检查通过。Excel 证据 `artifacts/e6/common-functions-excel-04`；桌面证据 `artifacts/avalonia/workspace/20261105-071655-51cafed2`。规范参考 [IFS](https://support.microsoft.com/en-gb/excel/functions/ifs-function)、[SUMPRODUCT](https://support.microsoft.com/en-us/excel/functions/sumproduct-function)。

发布与生产自检通过，活动包 `artifacts/avalonia/releases/20261105-071727-601471eb/application`。

## XC-0602c — 文本及数组函数 / Text and array formulas

2026-11-05，1.21.0。TEXTBEFORE/TEXTAFTER 支持正负序号、匹配大小写、将文本边界视为分隔符和未找到值；按搜索方向处理重叠分隔符。TEXTSPLIT 支持列/行分隔符、忽略空字段、匹配模式及补值；多个分隔符可通过矩形引用提供。SORTBY 支持多个单行或单列排序键、稳定顺序及正反向；TAKE/DROP 支持正负行列数、仅按列处理、截断到输入大小和空数组错误。

The six functions use the existing bounded array model, dependency tracking, spill ownership and XLSX function encoding. Text boundary functions accept scalar arguments. TEXTSPLIT accepts a scalar source and scalar options, with scalar or range delimiters. Array sorting and slicing preserve typed values and support composition with other array functions.

边界：TEXTBEFORE/TEXTAFTER 的数组参数、TEXTSPLIT 的数组源文本/控制参数、TAKE/DROP 的数组尺寸参数和 SORTBY 的数组方向参数尚未支持，导入时保持原有公式只读保留路径；数组字面量仍属于既有语法边界。数组结果最多 100,000 元素和 4 Mi 字符。排序沿用现有类型及不区分大小写文本规则，不声称新增地区排序能力。UTF-16 不拆分中文及非 BMP 文本的分隔字符。

39 项相关测试通过；追加存储/重开/增量缩小与撤销后 CommonFunctionTests 10 项通过。59 个累计真实 Excel 公式样例通过，来源 `artifacts/e6/common-functions-excel-05`；CommonFunctions 桌面检查 12 项通过，证据 `artifacts/avalonia/workspace/20261105-072315-25917d23`。规范参考 [TEXTBEFORE](https://support.microsoft.com/en-us/excel/functions/textbefore-function)、[TEXTSPLIT](https://support.microsoft.com/en-us/excel/functions/textsplit-function)、[TAKE](https://support.microsoft.com/en-us/excel/functions/take-function)。空源文本边界采用本机 Excel 16.0 build 20527 的独立样例（TEXTBEFORE 空源为 #N/A，TEXTSPLIT 空源为 #VALUE!），未以文档示意代替实测。

发布与生产自检通过，活动包 `artifacts/avalonia/releases/20261105-072720-269583df/application`。

## XC-0603 — 定位条件 / Go To Special

2026-11-05，1.22.0。编辑菜单新增定位条件，支持当前选区或当前表已用范围内的空白、常量、公式、错误值和可见单元格。空文本属于常量，返回空文本的公式属于公式，溢出子格不属于常量或空白；错误定位拒绝旧公式缓存。可见模式排除手动隐藏行列及筛除行。合并格以锚点分类，只选择完全在范围内、整体允许选择的合并区域；不扩展到范围外。

Go To Special creates a navigation-only selection. It deduplicates overlapping scopes, coalesces contiguous matches, respects selection protection and supports cancellation. Confirmation rechecks session, sheet, revision, calculated sheet identity and selection. Statistics, formatting and clear use the existing multi-area command paths; clear is one undoable operation.

范围：最多 64 个连续结果区域，空白/可见扫描范围最多 1,000,000 格，比较及访问预算 2,000,000；远端稀疏常量/公式无需枚举完整已用矩形。超限、取消和过期预览均不应用部分结果。后续写入继续受现有 20,000 格编辑预算、隐藏/筛选及保护规则约束。

验证：15 项定位、统计和本地化定向测试通过；8 项 Avalonia SpecialSelection 检查通过，包括取消、统计、多区格式、清除/一次撤销和过期预览。证据 `artifacts/avalonia/workspace/20261105-073642-55d889a4`。生产发布和自检通过，活动包 `artifacts/avalonia/releases/20261105-073652-32784661/application`，清单哈希已核对。

## XC-0604 — 可见行交换 / Visible-row transfer

2026-11-05，1.23.0。编辑菜单新增复制可见单元格和粘贴到可见单元格。跳过手动隐藏及筛除行，按源行顺序映射到目标可见行，目标必须显式选择同样的可见行列数。复制类型值及基本样式，支持单独只贴值；公式模式按每个源格原坐标转换到对应目标位置。目标隐藏/筛除行不写入。预览说明覆盖范围、记录数和跳过行数，整批通过 DocumentSession 提交并可一次撤销。

Visible-row transfers preserve typed identifiers, booleans, errors, dates and basic styles. A portable XML Spreadsheet payload supports another instance; local capture identity expires after any document or filter edit. Preview rechecks session, sheet, revision, calculated sheet identity and selection. Protection, validation and array ownership checks reject the whole transaction on failure.

映射范围：列位置保持不变，隐藏列须先取消隐藏；单次最多 20,000 可见格。富文本、链接、传统批注、源验证/Table 及未映射规则拒绝；目标验证保留并执行，值/公式粘贴保留目标条件格式，格式粘贴拒绝涉及条件格式。合并、浮动对象和线程讨论按整个包围矩形保守检查（包括隐藏行）。数组只可复制有效结果值，公式模式拒绝数组所有权。外部/内部公式统一采用现有外部协议的普通相对引用子集，拒绝名称、跨表、绝对/混合引用；XML 保存每格原始引用偏移。公式模式仅发布 XML，避免 HTML 压缩坐标下负偏移不能表达时悄悄更改语义，界面已提示。

验证：52 项可见行/外部导出/筛选剪贴板/本地化测试通过，涵盖类型、跨表和 XML 往返、日期系统、保存重开、保护与验证回滚、撤销/重做、取消和边界。9 项 Avalonia VisibleTransfer 检查通过，含公式路径、失效本地捕获及过期外部预览；证据 `artifacts/avalonia/workspace/20261105-074503-c289e027`。沿用 78 号 XML 编码协议，本批跨实例证据为序列化后去掉本地身份的接收路径，未将其称为新的实机 Excel 操作样例。

生产发布和自检通过，活动包 `artifacts/avalonia/releases/20261105-074519-26852aec/application`，活动清单哈希已核对。
