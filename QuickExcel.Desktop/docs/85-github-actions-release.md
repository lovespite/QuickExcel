# GitHub Actions Windows / macOS 发布

2026-11-20：4.2.0 接入 [release.yml](../../.github/workflows/release.yml)。构建脚本复用现有生产发布路径，Windows x64 与 macOS arm64 并行构建。Linux 不参与本工作流。

## 发布方式

1. 把发布改动提交并推送到 GitHub。产品版本取自 `QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj` 的 `Version`。
2. 在该提交创建同版本标签并推送：

   ```powershell
   git tag -a v4.2.0 -m "QuickExcel 4.2.0"
   git push origin v4.2.0
   ```

3. 在 GitHub Actions 查看 **Release Windows and macOS**。两端构建、自检、定向测试成功后，工作流自动创建 Release 草稿、上传两个平台的包和校验文件，核对 GitHub 的资产摘要，再公开 Release。

也可进入 Actions → Release Windows and macOS → Run workflow，填写**已存在**的标签。手动入口不会创建标签；工作流需先进入默认分支才会出现在该入口。只接受稳定版本 `vX.Y.Z`，拒绝版本不匹配、非标签引用、构建前后移动的标签。

无需个人访问令牌或自定义 Secrets。使用仓库的 `GITHUB_TOKEN`，校验/构建 job 只有 `contents: read`，最终发布 job 才有 `contents: write`。仓库或组织需要允许运行 GitHub Actions，以及工作流声明的 Release 写权限。

## 下载内容

| 资产 | 使用方式 |
| --- | --- |
| `QuickExcel-4.2.0-win-x64.zip` | 解压后运行 `application/quickxl.exe`；包含自包含 .NET 桌面与 Native AOT worker |
| `QuickExcel-4.2.0-osx-arm64.tar.gz` | Apple Silicon Mac 解压后执行 `QuickExcel/Start-QuickExcel.sh`；包含自包含 .NET、JSC 与 Seatbelt worker |
| 每个压缩包的 `.sha256` | 下载校验值及精确文件名 |

包内保留 `release.json`、依赖锁、许可证、JSC 对应源码与重建说明。Windows ZIP 当前不做 Authenticode 签名，也不包含 ClickOnce 安装器；目标设备的 Monaco/PDF.js 仍需要 WebView2 Evergreen Runtime。GitHub runner 在构建前安装/验证该运行时以执行生产自检。

macOS 沿用已经实现的原始目录分发路径，**没有 Developer ID 签名或 Apple 公证**；现有 `package-avalonia-macos.py` 的多文件 `.app` 签名限制尚未解决，因此本工作流不调用它。不提供 Intel Mac 包。持久脚本信任仍要求已签名的 bundle；当前包只允许显式单次授权。CI 自检不等于实体设备、签名、公证、IME、打印或完整 M6 验收。

## 构建与校验

- Windows 构建使用 `windows-2022`（保留 Visual Studio 2022 / v143 14.44），macOS 使用 `macos-26` arm64。SDK 在仓库 `artifacts/tools` 下独立安装为 .NET 10.0.404，避免 runner 自带的新主版本 SDK 参与；Node 22、Python 3.13。原生源版本和哈希继续由 `tools/QuickExcel.JscProbe` 的清单锁定。
- 显式安装 CMake 4.2.1 / Ninja 1.13.0，避免 Windows runner 自带 CMake 3.x 的差异。只跨运行缓存经哈希验证的上游下载归档；不缓存已修改的源目录、原生构建目录或生产包。原生构建默认并行度为 2，以适配标准 runner 内存。首次完整构建可能较慢，构建 job 上限为 350 分钟。
- Windows 使用现有 `build-avalonia-release.ps1 -SkipTests -RunSmoke`，独立发布 Native AOT worker，检查禁入产物及 Web 资源，执行生产自检，并将发布 worker 路径传入定向测试。CI 不修改本机活动包指针。
- macOS 构建 JSC 和原生 Seatbelt launcher/library，然后使用现有 Unix 发布脚本，执行生产自检和 macOS worker/策略相关测试，最后使用 `archive-unix-package.py` 保留可执行权限并验证归档内容。
- 两个平台从校验 job 解析出的相同 Git 提交构建。发布 job 再次验证标签、两个压缩包、校验文件及 GitHub 上传摘要。不会因某个平台失败而发布另一个平台的半成品。
- 所有 Actions 依赖固定到完整提交 SHA；更新时同时核对其版本注释和输入参数。

## 失败与重跑

构建失败时在 Actions 下载 `evidence-win-x64` / `evidence-osx-arm64`，查看原生编译日志、发布日志、生产自检结果和 TRX。证据保留 14 天，验证通过的下载暂存 artifact 保留 7 天。

上传中断会留下草稿，可重跑失败 job；也可手动填写相同标签重建。重跑只替换当前 workflow run 的暂存 artifact，以及尚未公开的 Release 草稿资产。草稿含无关附件时停止并要求维护者先检查。已经公开的版本禁止覆盖；需要变更时提升版本并推送新标签。

## 本地验证

```powershell
python scripts/test-github-release.py
# Windows 完整 CI 打包入口（会构建原生依赖及执行定向测试；需新输出目录）：
./scripts/build-github-windows-release.ps1
# 日常功能变更完成后仍按仓库约定更新本地启动包：
./scripts/build-quickexcel.ps1 -SkipTests -RunSmoke
```

macOS 上执行 `pwsh -NoProfile -File scripts/build-github-macos-release.ps1`。所有输出均在本仓库 `artifacts` 下。

本次验证记录在实现结束后补充；Windows 本地运行不能替代 GitHub 托管 runner 或 macOS 实机执行证据。

GitHub 平台依据：[runner 标签和资源](https://docs.github.com/en/actions/how-tos/write-workflows/choose-where-workflows-run/choose-the-runner-for-a-job)、[CLI Release 创建](https://cli.github.com/manual/gh_release_create)、[Release 资产 API](https://docs.github.com/en/rest/releases/assets)。
