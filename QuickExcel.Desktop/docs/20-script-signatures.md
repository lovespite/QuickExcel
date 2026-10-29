# 工作簿脚本数字签名与证书信任

2026-10-29。此功能更新此前文档中“打开一律不运行脚本”的桌面策略；未签名文件和恢复副本仍保持停用。存储层与 `WorkbookDocument.Open` 只读取文件，只有桌面完成签名验证与信任判断后才启用脚本。AppContainer、Job、IPC 和动作事务边界保持不变。

## 使用

在“脚本”页面左侧点击“签名…”，选择带私钥的 PFX/P12 代码签名证书，输入密码；也可“创建自签证书并签名…”，设置证书名称、私钥文件保存位置和至少 8 字符的 PFX 密码。新建证书为 RSA 3072，具有数字签名和代码签名用途，有效期两年。签名完成后保存 `.xlsq`。

| 状态 | 打开文件时的行为 |
| --- | --- |
| 未签名 | 保持停用；仍可点击“运行”手动授权 |
| 签名有效，Windows 代码签名证书链验证通过 | 自动启用、重算并发现操作脚本 |
| 签名有效，证书为未信任的自签证书 | 显示主体、颁发者、有效期、SHA-256 指纹，默认“取消” |
| 选择“仅本次允许” | 仅当前会话允许执行，不保存信任 |
| 选择“始终信任此证书” | 记住完整证书的 SHA-256 指纹，以后此证书签署的工作簿自动启用 |
| 内容或签名篡改、证书失效、用途不符、链失败 | 保持停用，不提供信任无效签名的选项 |

操作脚本只发现并列出，动作体仍需点击“执行动作”。模块初始化和 UDF 会在启用及重算时执行。信任证书意味着允许它签署的其他工作簿执行脚本，自签证书显示的名称不证明签署者身份。

“签名详情…”可查看验证结果、移除签名；“受信任证书…”可移除记住的自签证书，同时停用当前工作簿脚本。Windows 系统信任通过系统证书管理维护，QuickExcel 不向 Windows 根证书存储写入证书。

## 覆盖与失效

签名覆盖所有模块 ID、路径、原始源码（含库模块）、API 版本、入口顺序、全部 workbook 用途绑定，以及签名版本、算法和证书 SHA-256。模块和绑定 ordinal 排序，入口保持原序，源码换行及空白保持原样。

单元格数据、公式、工作表、工作簿身份、修订号和缓存不在签名范围内；数据编辑、重算、保存和另存保留签名。签名不能证明数据可信或脚本无害。

修改源码、模块或用途会移除签名并撤权，无效编辑器草稿也立即移除旧签名。单元格撤销/重做不恢复旧签名。签名变更经 `EditorController` → `DocumentSession` 提交，标记 dirty，不进入单元格撤销历史；提交要求模块仍与签名时捕获的对象相同。

## 文件与算法

XLSQ 保持格式 v2，新增必需能力 `script-signatures`，清单 `scriptSignature` 指向 `security/script-signature.json` 并带条目 SHA-256。旧读取器拒绝此能力；移除签名后保存删除能力、引用和条目。签名 JSON 上限 64 KiB，复用 ZIP/JSON 累计预算和重复属性检查。

JSON 为 `version: 1`、`algorithm`、Base64 DER `certificate` 和 Base64 `value`。支持 RSA-PSS-SHA256（至少 2048 位）与 ECDSA-P256-SHA256（P1363 固定长度）。私钥/密码不入包，PFX 私钥以 `EphemeralKeySet` 导入。

载荷用 .NET `BinaryWriter` 写入：严格 UTF-8，字符串长度为 7-bit 编码字节长度，Int32 小端。顺序为 `QuickExcel.ScriptSignature` 字符串、版本 Int32（1）、算法字符串、大写证书 SHA-256 字符串、API Int32、模块数量及排序后 `(ID,path,source)`、入口数量及入口序列、绑定数量及排序后 `(module,"workbook",kind)`。普通包摘要不能代替数字签名。

## 信任策略与限制

每次打开和点击“运行”重新验证。证书必须有代码签名 EKU，KeyUsage（若存在）允许 DigitalSignature，当前时间在有效期内。系统链使用代码签名策略、Online 撤销检查（排除根）、5 秒 URL 检索超时；撤销未知/离线等失败不自动信任。参考 [.NET X509ChainPolicy](https://learn.microsoft.com/dotnet/api/System.Security.Cryptography.X509Certificates.X509ChainPolicy)。

只有链仅含该证书、subject/issuer 一致且唯一错误为 UntrustedRoot 的自签证书可以被记住。指纹不覆盖过期、撤销、显式不信任、用途或其他链错误。同名、换密钥或重新签发均不是原证书。无时间戳服务，过期后需重新签名；自签证书无在线撤销服务，应移除不再信任的证书。

信任列表 `%LOCALAPPDATA%/QuickExcel/trusted-script-certificates.json` 只保存指纹、显示名称和到期时间；跨进程独占文件锁和原子替换保护写入。读取损坏或写入失败不授予授权。列表属于当前 Windows 用户，不防御同账户恶意进程修改。测试及 `--ui-validation` 使用指定目录中的独立列表。

文件不保存“已信任”标志。恢复入口不自动执行，用户审阅后可手动运行。`.xlsx` 导出不含脚本及签名。此功能不等同于程序/安装器 Authenticode 签名或完成 M6 OS 门禁。

## 验证

按修改范围运行，未执行全量回归：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --no-build --artifacts-path .\artifacts\build -v minimal --filter "FullyQualifiedName~ScriptSignatureTests|FullyQualifiedName~TypedStorageTests|FullyQualifiedName~DocumentSessionTests|FullyQualifiedName~M3StorageTests|FullyQualifiedName~WorkbookFacadeTests"
.\scripts\test-script-signatures.ps1
```

Release 构建 0 警告/0 错误。正常用户环境相关测试 125/125（签名 26 项），结果 `artifacts/validation/script-signatures/related-tests-final.trx`。受限环境首轮 105/106，PFX 导入因环境权限失败；正常用户测试通过，未更改生产权限要求。

WPF 专项覆盖真实信任按钮、默认取消、一次/持久信任、重开自动授权、动作仅发现、数据编辑保留签名、源码编辑撤权、篡改拒绝、损坏信任列表撤权、恢复入口和撤销信任。公式和动作发现使用实际 AppContainer worker。截图、逐项结果和二进制哈希位于测试脚本输出的唯一 artifacts 目录；程序化断言不代表真实鼠标/键盘或完整 OS 验收。
