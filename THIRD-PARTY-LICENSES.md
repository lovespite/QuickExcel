# 第三方组件与许可

本项目包含以下第三方材料。各组件的许可证仅覆盖其自身内容，不改变 QuickExcel 其余代码的许可。

| 组件 | 许可证 | 位置 | 说明 |
| --- | --- | --- | --- |
| RichTextKit | Apache-2.0 | `third_party/RichTextKit` | 文本布局、换行与 HarfBuzz 集成；保留原始版权头与 LICENSE |
| Office Scripts API 定义 | CC BY 4.0 | `third_party/office-scripts-docs` | 仅用于生成脚本类型声明；QuickExcel 自行实现全部行为 |

## Office Scripts API 定义

来源为 Microsoft 的 `office-scripts-docs-reference` 仓库，固定 commit 与校验值记录在 `third_party/office-scripts-docs/source.lock.json`。生成物改写了命名空间、删除了文档注释并追加了 QuickExcel 扩展，属于修改后的材料。

分发包含该声明或其派生声明的产物时，必须同时保留：

- 版权声明与来源链接；
- CC BY 4.0 许可证链接；
- 已修改的说明。

完整声明见 [third_party/office-scripts-docs/NOTICE.md](third_party/office-scripts-docs/NOTICE.md)。

## RichTextKit

来源、修改说明与 SHA-256 见 [third_party/RichTextKit/SOURCES.md](third_party/RichTextKit/SOURCES.md)。
