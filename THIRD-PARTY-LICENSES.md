# 第三方组件与许可

本项目包含以下第三方材料。各组件的许可证仅覆盖其自身内容，不改变 QuickExcel 其余代码的许可。

| 组件 | 许可证 | 位置 | 说明 |
| --- | --- | --- | --- |
| RichTextKit | Apache-2.0 | `third_party/RichTextKit` | 文本布局、换行与 HarfBuzz 集成；保留原始版权头与 LICENSE |
| Office Scripts API 定义 | CC BY 4.0 | `third_party/office-scripts-docs` | 仅用于生成脚本类型声明；QuickExcel 自行实现全部行为 |
| Avalonia、Avalonia WebView、BuildServices、MicroCom、Tmds.DBus、Fluent System Icons | MIT | Licenses | 桌面界面、平台集成和图标 |
| FastWpfGrid | MIT | licenses/QuickExcel.Grid.Shared | 保留的网格算法与来源声明 |
| ANGLE | BSD 及随附声明 | Licenses/ANGLE-LICENSE.txt | 图形后端 |
| SkiaSharp、HarfBuzzSharp | MIT 及原生库各自条款 | SKIA 与 HARFBUZZ 的 LICENSE 和 THIRD-PARTY-NOTICES | 图形与字体 |
| LibTiff.NET、PdfPig | 分别为 BSD-3-Clause、Apache-2.0 | LIBTIFF-LICENSE.txt、PDFPIG-LICENSE.txt | 图像和 PDF |
| AngleSharp、OpenMcdf | 分别为 MIT、MPL-2.0 | Licenses/ANGLESHARP-LICENSE.txt、Licenses/OPENMCDF-LICENSE.txt | HTML 和复合文档读取 |
| Monaco Editor、PDF.js、esbuild | 分别为 MIT、Apache-2.0、MIT | Web/licenses | 编辑器、PDF 预览与构建；字体、字符映射和 WASM 附各自许可 |
| JavaScriptCore / WebKit、ICU、mimalloc、TypeScript | 各组件分别授权 | JSC-Licenses | 脚本包附版权与许可原文、对应源码归档及重建说明；Windows VC++ 运行库适用 Microsoft 条款 |
| .NET 运行时 | MIT 及随附第三方条款 | 自包含发行包的 DOTNET-LICENSE 和 DOTNET-THIRD-PARTY-NOTICES | 桌面及隔离进程 |

## Office Scripts API 定义

来源为 Microsoft 的 `office-scripts-docs-reference` 仓库，固定 commit 与校验值记录在 `third_party/office-scripts-docs/source.lock.json`。生成物改写了命名空间、删除了文档注释并追加了 QuickExcel 扩展，属于修改后的材料。

分发包含该声明或其派生声明的产物时，必须同时保留：

- 版权声明与来源链接；
- CC BY 4.0 许可证链接；
- 已修改的说明。

完整声明见 [third_party/office-scripts-docs/NOTICE.md](third_party/office-scripts-docs/NOTICE.md)。

## RichTextKit

来源、修改说明与 SHA-256 见 [third_party/RichTextKit/SOURCES.md](third_party/RichTextKit/SOURCES.md)。
