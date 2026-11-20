# XC-0003：兼容诊断定位收尾

2026-10-30 完成 XC-0003，提交为本记录所在的 `feat: complete located compatibility diagnostics (XC-0003)`。本记录更新 [28](28-xlsx-compatibility-diagnostics.md) 的“未完成范围”，并结合 [37](37-excel-business-baseline.md)/[38](38-excel-edit-matrix.md) 关闭 E0 声明范围。E1 七项已有各自真实 Excel 对照，下一项为 XC-0201；M4/M6、发布及 E2/E3 仍未完成。

## 五类报告与来源位置

可编辑、仅保留、内置不可计算/待验证、潜在损失和拒绝分组保持。新增定位，不修改文件特性支持范围或保存策略：

- 共享字符串问题保留 `sharedStrings[i]` 来源索引，并按实际共享索引展开到受影响工作表/单元格。
- 共享字体、填充、边框、对齐问题保留 `cellXfs[i]`；命名基样式问题沿 `cellStyleXfs[i] → cellXfs[j]` 关联展开。显式单格与行/列继承使用实际有效样式，整行/列格式用 `row:n` / `columns:min:max` 表示，不枚举空白工作表。
- 来源位置与受影响位置均保留，特性列表仍是**导入时**定位，改名后不伪称当前定位。编辑后转为潜在损失，撤销恢复一致时回到仅保留；XLSQ 重开重建同样记录。
- 实际重算后的 Blocked 依赖现在定位到当前格，旧缓存不标新鲜。报告不执行计算或脚本；未重算时保留直接函数/语法定位和待验证说明，不凭缓存猜测整个计算图已成功。

共享来源索引与定位总量受现有 1024 位置上限约束，超出明确省略；独立损失说明的拒绝预算不放宽。新增定位复用同一损失说明，不因共享字符串被引用 1030 次就当成“1030 种问题”，也不因位置省略而允许有损保存。无共享问题时不创建逐格展开记录。

## 拒绝与界面

导入失败在原异常附加结构化信息，保留异常类型与调用方契约。能确定时记录 XML 部件、工作表和单元格/行列范围；工作簿级错误不借用最后读取的单元格。ZIP/文件或无法进一步定位的错误只标文件层级，不猜测范围。取消仍是取消。

桌面实际打开错误显示 `【拒绝】工作表 · 地址 · 部件 · 特性：原因`，失败不替换当前文档。比较工具的导入/编辑拒绝进入 Rejected 分组；编辑拒绝保留命令和来源首表名称。普通交互命令仍通过既有原子命令失败入口说明原因，不把失败尝试持久化成工作簿内容。

报告文案已同步 XLSX/XLSQ 原包保存能力。已支持特性不因旧笼统分类重新加损失提示；未知修订/拼音/对象等警告继续保留。用户仍须明确选择受支持副本，报告不会静默批准删除未知内容。

## 证据

只跑相关存储、诊断、公式保留、内容及 WPF 专项，未运行全量套件或发布。路径均在 `artifacts/compatibility` 下。

| 检查 | 结果 / 证据 |
| --- | --- |
| Core 诊断、typed 存储、受控公式、富文本相关 | 112/112，`xc0003-related`；含实际 Blocked 依赖、异常类型不变、只读及缓存保护 |
| 共享/既有位置及保存保护专项 | 9/9，`xc0003-shared-final`；命名基样式追加 1/1，`xc0003-named-final` |
| 最终位置专项 | 6/6，`xc0003-location-verified`；共享复用、1030 引用上限、命名关联、无效样式/字符串索引及工作簿级拒绝不误标单元格 |
| WPF 实际窗口/保存/打开处理器 | 103/103，`xc0003-closure-wpf-01`；新增 6 项：A1/B2 展开、报告只读、损失分组、拒绝副本保存、撤销恢复、损坏文件定位并保留当前文档 |
| 合成诊断基线 | `xc0003-closure-baseline`：12 来源哈希保持，96 路径中 72 输出 schema 0 问题、24 命令拒绝、8 普通保存拒绝后显式诊断副本；不称为 72 次 Excel 验收 |
| 真实 Excel 来源复验 | `xc0003-excel-import-verified/business-baseline-report.json`：复用 Excel 16.0 build 20527、语言 2052、自动计算的六份来源及固定期望，5384 属性通过，原包字节不变、schema 0 问题；实际 Excel 重开/重算证据见 37/38 |

WPF 文字证据为 `shared-style-report.txt` 和 `located-rejection-log.txt`；不是物理输入、OS IME、跨 DPI 或性能门禁。最终位置层级修正与无共享警告快速路径不改 UI 行为，分别由位置专项和真实 Excel 来源复验覆盖。Desktop Release 构建通过，仅既有 NU1900 警告。

初次测试的名称空间构造和 ImmutableArray 断言失败保留在 `xc0003-shared-01/-02/-03`，最终通过记录另存。E0 关闭指当前 E1 所需样本、比较与报告能力具备；未知 Excel 特性仍可能只保留或拒绝，阶段完成不表示任意 XLSX 无损编辑。

## 2026-11-20：Excel 附加元数据保留与按类型统计

本节更新以上旧的定位上限和“未知修订/拼音”边界。`artifacts/test/demo.xlsx` 原有 12 类兼容说明中，1307 行重复的 `x14ac:dyDescent` 占满了 1024 个定位样本；旧界面将样本和容量提示累计显示为 1025 处。

新增保留范围：

| 位置 | 内容与保存方式 |
| --- | --- |
| workbook/fileVersion | 保留已验证的生产程序版本属性 |
| workbook/mc:AlternateContent | 仅保留单个 x15 Choice 中的 x15ac:absPath；路径仅作为字符串，不访问目标 |
| workbook/xr:revisionPtr | 保留已验证的修订/协作标识；不赋予运行语义 |
| workbook/@mc:Ignorable、worksheet/@mc:Ignorable | 同时保留命名空间绑定，不能用 Ignorable 隐藏实际未知结构 |
| workbookPr/@codeName | 保留工作簿代码名称；不启用宏或脚本 |
| worksheet/phoneticPr | 保存类型、对齐与独立字体载荷；写出时复用或追加字体并重映射 fontId |
| worksheet/ignoredErrors | 建模明确范围及九种错误提示标志，保持规则顺序 |
| worksheet/@xr:uid、workbookView/@xr2:uid | 原对象保留标识；工作表按稳定 SheetId 对应原部件，复制表不复用原 uid |
| row/@x14ac:dyDescent、sheetFormatPr/@x14ac:dyDescent | 独立保存有限、非负的原值；行属性随行移动，输出补齐正确的 Ignorable 声明 |

普通单元格编辑可保留上述内容，不再要求有损副本。未知分支、未知属性、非法结构仍走原有诊断和保存保护。`ignoredErrors` 支持行列插删、移动、剪切、完整粘贴和转置的范围映射；涉及规则的自动填充、整行列复制以及不能安全映射的部分范围排序仍显式拒绝。值/公式/格式粘贴保留目标规则。拼音显示设置与默认行排版信息不阻止分类汇总。

新写 XLSQ 声明 `worksheet-excel-metadata` 必需能力，避免旧阅读器在编辑后遗漏这些字段；新阅读器仍接受历史 XLSQ v2 中尚未声明该能力的原始 XLSX 元数据。字体 XML 只作为有预算、禁止 DTD 的惰性值处理，不进入脚本或外部资源加载路径。

报告保留五种能力分组，但组内按语义类型统计；同类型在多行、多格出现仍为一类，受影响行数按工作表和行区间去重后单独显示。部件、整列等不能确定行数的位置明确标注，不编造行数。1024 上限只限制详细位置样本，不再截断后续类型与行统计；容量提示不计入类型数。聚合类型、行区间和共享来源均有独立资源预算，超限明确拒绝，不显示伪完整的统计。

定向验证入口为 `OfficeMetadataPreservationTests`、`ExcelMetadataCommandsTests`、`CompatibilityAggregationTests`、`CompatibilityReportTests`、CompatibilityProbe 的 `OfficeMetadataSchemaTests` 与共享位置测试，以及 `scripts/test-avalonia-workspace.ps1 -Mode Compatibility`。真实 demo 的原文件哈希、XLSX/XLSQ 编辑往返证据保存在 `artifacts/investigations/demo-compatibility-20261120-123441/implemented/verification.json`；该证据不等同于 Excel 物理操作验收。

最终定向检查为 Core 236/236、OpenXML/共享定位 12/12、实际 Avalonia 兼容弹窗 37/37。测试记录位于 `artifacts/test/office-metadata/*-verified.trx`，弹窗证据位于 `artifacts/avalonia/workspace/20261120-045009-5cfbd283`。`build-quickexcel.ps1 -SkipTests -RunSmoke` 发布及生产自检通过；当前启动指针选中 `20261120-045232-7d56f86a/application`，其 Core、Infrastructure、Presentation、Desktop 四个程序集与本次验证发布 `20261120-045216-59f24a63/application` 的 SHA-256 完全一致，指针的 manifest 哈希与生产自检状态均已核验。
