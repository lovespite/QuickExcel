# 命名样式、主题字体和扩展对齐

2026-10-30，XC-0102 的常用读显写范围完成；单格编辑和既有行/列优先级可用。此前行样式切片已在 `65309c0`，本轮实现由本文件首次提交定位。未新增命名样式管理或主题编辑界面。

## 契约

- 校验 Normal/命名样式名称、`xfId`、基表组件索引和数字格式引用。有效显示以当前 `cellXfs` 的组件为准；Excel 的 `apply*` 还控制父样式变化后的传播，不能把 false 直接解释为忽略当前组件。参见 [Microsoft XF 实现说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/59922f8b-0edc-4e93-a822-9f22254aec46)。
- 重写时保留原样式表中现存 XF、基样式、命名样式、颜色及主题引用的原始索引。仅为修改后的样式追加组件和 XF；不修改其他单元格的共享样式。新样式保留命名关联，并将格式分量显式覆盖；它不是在 QuickExcel 内修改命名样式定义。
- 有范围限制的样式部件保留与无外部依赖的主题保留接入存储。样式默认切片器/时间线样式名与主题家族标识按确切结构白名单保留；未知扩展或关系仍报告。无法证明命名关联安全时，即使选择受支持副本也拒绝写入，原目标保持不变。
- 解析主题 major/minor 拉丁字体，东亚字体使用显式 ea 或 Hans 回退，WPF 与编辑器使用本机字体回退列表。显式选择新字体清除主题方案关联。主题其余静态数据保留，不加载外部关系中的字体资源。
- 增加旋转 0–180（对应 Excel -90°–90°）、缩进 0–250、非换行文本缩小填充的读显写；旋转进入自动行列尺寸测量。缩小后重新测量实际字形，避免仍出现省略号。缩小尺寸有 1 DIP 下限和有界迭代。
- 行/列与单格样式优先级、`customFormat`、空白格录入/清除仍沿用 [29](29-xlsx-row-style-inheritance.md)。没有升级 XLSQ：新字段已有标准 XLSX 表达，旧读取器会按原有未知特性保护处理。

历史边界（2026-10-30）：255 竖排文字当时仅保留并报告，2026-11-05 的新增实现见下文；复杂旋转/换行组合、非 Hans 文字的主题脚本选择与完整排版没有声称等同 Excel。斜线边框、花纹填充、字体上下标等原有显示缺口仍有诊断。未取得来源关联的人工构建非零 NamedStyleId 不可凭空写出；未知主题/样式依赖不得通过复制未知 ZIP 条目绕过。

## 验证

本机 Excel 16.0 build 20527、64 位、语言 2052、1900 日期系统；所有 Excel 输入均为脚本新建的隔离夹具。数字格式上一项提交为 `19b11e3`，本轮复验其来源导入与重写 86 项仍通过。

| 验证 | 结果与证据 |
| --- | --- |
| 命名样式、主题字体、行继承、M3 存储 | `xc0102-style-tests` 中 13:49:15 的 17 项通过；最后增加的命名样式安全边界 13:51:20，4/4 |
| 既有类型存储/报告 | 首轮 93 项中 79 项通过；14 项旧主题损失预期随新增保留能力更新，13:44:08 的同组 14 项通过；失败 TRX 保留 |
| 比较工具 | 34/34，`artifacts/compatibility/xc0102-style-probe-tests/ASUS_LOVESPITE-PC_2026-10-30_13_49_18_net10.0.trx` |
| Excel 生成源 | `artifacts/compatibility/xc0102-excel-02/`：主题 major/minor、命名金额样式、旋转、缩进、缩小填充 |
| 改 B1 值与 A3 斜体后写回 | `artifacts/compatibility/xc0102-style-roundtrip-02/styles-report.json`：原 XF/命名表/主题保持，源哈希未变，独立 schema 0 错误 |
| Excel 正常只读重开 | `artifacts/compatibility/xc0102-excel-reopen-02/comparison.json`，8 格 × 13 属性，104/104；只允许指定 A3 Italic 改变 |
| WPF 像素、尺寸与既有渲染 | `artifacts/compatibility/xc0102-styles-wpf-02/`，36/36；截图已检查 |
| Release | Desktop 构建通过；NU1900 漏洞源不可达警告仍在 |

Excel `.Style` 在此 COM 环境返回 Style 对象；最终脚本记录其 Name，早期 `-01` 的空对象不算名称验收。初版缩小填充在截图上仍有省略号，最终 `-02` 已修正并增加完整行宽断言。原失败证据未覆盖。

剩余工作簿视图、页面与文档属性仍属 XC-0103/0104/0105；真实 Excel 来源的本轮输出明确命名为 `diagnostic-copy`，不是全文件无损保存结论。报价合成夹具的全部八条编辑路径已不再触发原样式继承损失提示，并保留独立结构检查。

复现：

```powershell
.\scripts\test-excel-styles.ps1 -OutputDirectory artifacts/compatibility/new-styles
dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- styles artifacts/compatibility/new-styles/excel-source.xlsx artifacts/compatibility/new-styles-roundtrip
.\scripts\test-excel-styles.ps1 -OutputDirectory artifacts/compatibility/new-styles-reopen -WorkbookPath artifacts/compatibility/new-styles-roundtrip/edited.diagnostic-copy.xlsx
```

仅运行相关测试和 WPF 专项，未跑全套回归、基准或发布；managed 命令均指定 artifacts 路径，未修改同级 ExcelIO。


## 2026-11-05：单元格文字方向

开始 → 对齐 → 文字方向提供逆时针 45°、顺时针 45°、竖排文字、向上 90°、向下 90°和水平文字。菜单显示当前单元格的方向；“设置单元格对齐方式…”直接打开对齐页，可输入 −90° 到 90° 的整数角度，或选择竖排并查看预览。正角度为逆时针，负角度为顺时针。

设置沿用选区格式命令、可见行筛选规则及原有编辑提交门禁，一次撤销/重做；不修改值、公式或富文本内容。编辑框保持普通文字方向，合并单元格仍由锚点持有编辑草稿。XLSX/XLSQ 和已有剪贴板编码继续使用 `textRotation=0..180/255`，不升级文件格式。

Skia 网格、打印和自动尺寸测量共用竖排布局：每个 Unicode 文本元素保持正向并按列中心对齐，保留分段字体、颜色和装饰。`255` 已可显示，读取时不再报告“竖排文字暂仅保留”或因此阻止正常导出。复杂旋转/自动换行组合及东亚标点、跨列竖排排版未声称与 Excel 完全相同。

定向验证：`TextRotationTests` 与剪贴板旋转往返共 16 项通过；渲染相关 76 项通过，覆盖方向、居中、裁剪、组合字符及富文本字形；`scripts/test-avalonia-workspace.ps1 -Mode TextRotation` 的 37 项检查通过，覆盖菜单、范围应用、普通编辑框像素、合并锚点、自定义角度、取消及保存重开；证据为 `artifacts/avalonia/workspace/20261105-041902-bf5642f7`。数字格式对话框回归 30 项通过，证据为 `artifacts/avalonia/workspace/20261105-041932-d6c58336`。仅运行相关检查，不运行完整回归套件。本地自包含包按仓库要求通过 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 发布和激活。
