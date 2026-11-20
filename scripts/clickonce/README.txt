QuickExcel Windows x64 — ClickOnce 安装包

1. 解压完整 ZIP，保留全部目录和文件，再运行 setup.exe。
2. setup.exe 检测 WebView2 Runtime；缺少时使用包内微软在线引导程序下载并安装，
   此步骤需要连接 Internet。已经安装运行库的设备跳过此步骤。
3. 在 ClickOnce 提示中核对发布者及版本并安装。安装后可从开始菜单启动。

应用包含部分裁剪的 .NET 10；隔离脚本 worker 使用 Native AOT，原生依赖随包携带。
JSC、ICU、TypeScript 只保留在 ScriptWorker 目录，无需另装 .NET SDK 或运行时。
系统要求：受支持的 Windows 10/11 x64，启用系统自带的 .NET Framework 4.8
和 ClickOnce。包内不再携带 WebView2 完整离线安装器。请勿只复制 setup.exe。
本包不配置固定发布地址或自动更新；后续可从新版本的完整包手动安装更新。

签名采用 QuickExcel 自签代码签名证书。自签证书不自动获得 Windows 信任，
因此可能出现“未知发布者”或策略阻止。请先通过可信渠道核对随包公钥证书
QuickExcel-CodeSigning.cer 的指纹（记录于 clickonce-release.json）。
如设备策略要求受信任发布者，由用户/管理员按组织策略显式导入公钥证书至
当前用户的“受信任的根证书颁发机构”和“受信任的发布者”。安装器不自动
修改信任存储；不要导入来历不明的证书。此包不包含 PFX 私钥或其密码。

发布版本自动来自项目内部版本；ClickOnce 使用等价的四段数字版本。
clickonce-release.json 记录版本、证书、组件来源及逐文件 SHA-256。
应用“设置 → 关于”中可查看第三方声明及完整开源许可证。
ClickOnce 不内置 JSC 源码压缩包；JSC-Licenses/JSC-SOURCE.txt 提供实际使用的
完整源码地址（含 QuickExcel 修改和构建脚本），也可在设置的许可证列表中查看。
Windows 主程序文件为 quickxl.exe；产品、托管程序集和隔离 worker 名称保持不变。
