[CmdletBinding()]
param([switch]$SkipTests, [switch]$CreateArchive, [switch]$RunSmoke,
      [string]$OutputDirectory = '', [switch]$Activate,
      [ValidateSet('Jsc')][string]$ScriptEngine = 'Jsc',
      [string]$CodeSigningCertificateThumbprint = '')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$previousBuildPath = $env:PATH
$previousDotnetRoot = $env:DOTNET_ROOT
$previousBuildOs = $env:OS
try {
# The servicing linker runs on the matching servicing runtime. A repository-local
# official SDK can be used without replacing the machine's SDK installation.
$localSdk = Join-Path $root 'artifacts/tools/dotnet-10.0.404'
if (Test-Path -LiteralPath (Join-Path $localSdk 'dotnet.exe')) {
    $env:PATH = $localSdk + [IO.Path]::PathSeparator + $env:PATH
    $env:DOTNET_ROOT = $localSdk
}
if ([version](& dotnet --version) -lt [version]'10.0.404') { throw 'Windows releases require .NET SDK 10.0.404 or newer. An official SDK ZIP may be extracted under artifacts/tools/dotnet-10.0.404.' }
# Some task runners omit this standard Windows variable; Native AOT's targets
# require it even when MSBuild has already identified the real host platform.
if ($IsWindows -and -not $env:OS) { $env:OS = 'Windows_NT' }
$version=([xml](Get-Content -LiteralPath (Join-Path $root 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj') -Raw)).SelectSingleNode('/Project/PropertyGroup/Version').InnerText
$artifacts = Join-Path $root 'artifacts'
$runId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8)
function ArtifactPath([string]$value) {
    $path = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($value)) { $value } else { Join-Path $root $value }))
    if (-not $path.StartsWith($artifacts + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Output must stay under this repository artifacts directory.' }
    $parent = $path
    while ($parent.Length -ge $artifacts.Length) {
        if ((Test-Path -LiteralPath $parent) -and ((Get-Item -LiteralPath $parent -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Output cannot traverse reparse points: $parent" }
        $parent = Split-Path $parent -Parent
    }
    $path
}
$package = ArtifactPath $(if ($OutputDirectory) { $OutputDirectory } else { "artifacts/avalonia/releases/$runId/application" })
if (Test-Path -LiteralPath $package) { throw 'A release must use a new directory; previous packages are retained.' }
$evidence = ArtifactPath "artifacts/avalonia/releases/$runId/evidence"
$build = ArtifactPath "artifacts/avalonia/release-build-$ScriptEngine"
$lockDirectory = Join-Path $evidence 'nuget-locks'
$workerBuild = ArtifactPath "artifacts/avalonia/release-build-$ScriptEngine-worker-aot"
$workerLockDirectory = Join-Path $evidence 'worker-nuget-locks'
$workerDirectory = Join-Path $evidence 'worker-native'
$project = 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj'
New-Item -ItemType Directory -Path $package,$evidence,$lockDirectory,$workerLockDirectory -Force | Out-Null
Push-Location $root
try {
    & "$PSScriptRoot/build-avalonia-web.ps1"
    if (-not $SkipTests) {
        dotnet test QuickExcel.Core.Test/QuickExcel.Core.Test.csproj -c Release --artifacts-path $build "-p:QuickExcelScriptEngine=$ScriptEngine" --filter 'FullyQualifiedName~DesktopSettingsMigrationTests|FullyQualifiedName~WebDraftContractTests|FullyQualifiedName~LocalWebOriginTests|FullyQualifiedName~M4IsolationTests|FullyQualifiedName~M4WorkerTests|FullyQualifiedName~M4WorkerFaultTests|FullyQualifiedName~M4RecoveryTests|FullyQualifiedName~TypeScriptCompilerTests|FullyQualifiedName~TypeScriptStorageTests|FullyQualifiedName~JscBridgeTests' --logger 'trx;LogFileName=release-contracts.trx' --results-directory $evidence -v minimal 2>&1 | Tee-Object (Join-Path $evidence 'tests.log')
        if ($LASTEXITCODE -ne 0) { throw 'Release contract/isolation checks failed.' }
    }
    # RID restore is explicit; resolved versions and all source locks enter package identity.
    foreach ($lock in @(& git ls-files '*packages.lock.json')) {
        $projects=@(Get-ChildItem -LiteralPath (Split-Path (Join-Path $root $lock) -Parent) -Filter '*.csproj')
        if ($projects.Count -eq 1) { Copy-Item -LiteralPath (Join-Path $root $lock) -Destination (Join-Path $lockDirectory ($projects[0].BaseName+'.packages.lock.json')) }
    }
    dotnet restore $project -r win-x64 -p:RuntimeIdentifier=win-x64 --artifacts-path $build "-p:QuickExcelScriptEngine=$ScriptEngine" -p:QuickExcelDiagnostics=false -p:SelfContained=true -p:RestoreForceEvaluate=true "-p:QuickExcelReleaseLockDirectory=$lockDirectory" -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'RID restore failed.' }
    # Project-reference restores can drop a referenced executable's RID. Keep
    # the independent AOT dependency graph/cache/locks separate from the desktop.
    dotnet restore QuickExcel.ScriptWorker/QuickExcel.ScriptWorker.csproj -r win-x64 -p:RuntimeIdentifier=win-x64 --artifacts-path $workerBuild -p:SelfContained=true -p:RestoreForceEvaluate=true "-p:QuickExcelReleaseLockDirectory=$workerLockDirectory" -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'Native AOT worker restore failed.' }
    dotnet publish QuickExcel.ScriptWorker/QuickExcel.ScriptWorker.csproj -c Release -r win-x64 --self-contained true --artifacts-path $workerBuild -p:RestoreLockedMode=true "-p:QuickExcelReleaseLockDirectory=$workerLockDirectory" -p:ILLinkTreatWarningsAsErrors=true -p:DebugType=None -p:DebugSymbols=false -o $workerDirectory -v minimal 2>&1 | Tee-Object (Join-Path $evidence 'worker-publish.log')
    if ($LASTEXITCODE -ne 0) { throw 'Native AOT worker publish failed.' }
    dotnet publish $project -c Release -r win-x64 --self-contained true --artifacts-path $build "-p:QuickExcelScriptEngine=$ScriptEngine" -p:QuickExcelDiagnostics=false -p:RestoreLockedMode=true "-p:QuickExcelReleaseLockDirectory=$lockDirectory" "-p:QuickExcelPublishedWorkerDirectory=$workerDirectory" -p:DebugType=None -p:DebugSymbols=false -o $package -v minimal 2>&1 | Tee-Object (Join-Path $evidence 'publish.log')
    if ($LASTEXITCODE -ne 0) { throw 'Avalonia release publish failed.' }
    $engineFiles = @('JavaScriptCore.dll','icudt77.dll','icuuc77.dll','icuin77.dll','msvcp140.dll','msvcp140_2.dll','vcruntime140.dll','vcruntime140_1.dll','typescript.bundle.json','JSC-BUILD-IDENTITY.json')
    foreach ($notice in @('THIRD-PARTY-LICENSES.md','OFFICE-SCRIPTS-NOTICE.md','MARKDOWN-AVALONIA-LICENSE.txt','DOTNET-LICENSE.txt','DOTNET-THIRD-PARTY-NOTICES.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $package "Licenses/$notice") -PathType Leaf)) { throw "Missing product license notice: $notice" }
    }
    foreach ($file in (@('quickxl.exe','QuickExcel.Avalonia.dll','QuickExcel.Avalonia.deps.json','Assets/QuickExcel-XLSQ.ico','Assets/QuickExcel-XLSX.ico','Assets/QuickExcel-CSV.ico','coreclr.dll','hostfxr.dll','QuickExcel.Grid.Avalonia.dll','libSkiaSharp.dll','libHarfBuzzSharp.dll','Web/manifest.json','ScriptWorker/QuickExcel.ScriptWorker.exe') + @($engineFiles | ForEach-Object { 'ScriptWorker/' + $_ }))) {
        if (-not (Test-Path -LiteralPath (Join-Path $package $file) -PathType Leaf)) { throw "Missing release payload: $file" }
    }
    & "$PSScriptRoot/test-windows-payload.ps1" -PackageDirectory $package
    if (Test-Path -LiteralPath (Join-Path $package 'QuickExcel.Avalonia.exe')) { throw 'The Windows package must contain only the renamed quickxl.exe apphost.' }
    foreach ($file in @('Markdown.Avalonia.dll','ColorDocument.Avalonia.dll','ColorTextBlock.Avalonia.dll')) {
        if (-not (Test-Path -LiteralPath (Join-Path $package $file) -PathType Leaf)) { throw "Missing license viewer dependency: $file" }
    }
    $all = @(Get-ChildItem -LiteralPath $package -File -Recurse)
    & python (Join-Path $root 'tools/QuickExcel.JscProbe/prepare_distribution.py')
    if ($LASTEXITCODE -ne 0) { throw 'JSC source/notices preparation failed.' }
    & {
        if ($all | Where-Object { $_.Name -match 'quickexcel-js|QUICKJS-LICENSE' }) { throw 'Legacy engine entered the release.' }
        $licenses = Join-Path $package 'JSC-Licenses'; New-Item -ItemType Directory -Path $licenses | Out-Null
        foreach ($name in @('JSC-THIRD-PARTY-NOTICES.txt','JSC-CORRESPONDING-SOURCE.zip','JSC-REBUILD.txt')) {
            Copy-Item -LiteralPath (Join-Path $root "artifacts/jsc/distribution/$name") -Destination $licenses
        }
    }
    $forbidden = @($all | Where-Object { $_.Name -match 'RecoveryFixture|WorkerFixture|HostFixture|SandboxProbe|PlatformSmoke|TestProbe|TypeScriptProbe|JscProbe|quickexcel-jsc-probe|Designer|Avalonia\.Diagnostics|\.pdb$|\.map$|\.pfx$|ExcelIO|Grid\.Wpf|^PresentationFramework|^PresentationCore|Xpf|quickjs-eval' })
    if ($forbidden) { throw "Forbidden release payload: $($forbidden.FullName -join ', ')" }
    if (@($all | Where-Object { $_.FullName.StartsWith((Join-Path $package 'ScriptWorker') + [IO.Path]::DirectorySeparatorChar) -and $_.Name -match 'Avalonia|WebView|Skia|HarfBuzz|RichTextKit|PdfPig|Rendering|LibTiff|Monaco' })) { throw 'UI/Web/rendering payload leaked into ScriptWorker.' }
    $web = Get-Content -LiteralPath (Join-Path $package 'Web/manifest.json') -Raw | ConvertFrom-Json
    if ($web.testBuild -or $web.probeBuild) { throw 'Release contains test/probe Web resources.' }
    foreach ($asset in $web.files) {
        $file = Join-Path $package ('Web/' + $asset.name)
        if ((Get-Item -LiteralPath $file).Length -ne $asset.size -or (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne $asset.sha256) { throw "Invalid Web resource: $($asset.name)" }
    }
    if (@(Get-ChildItem (Join-Path $package 'Web') -File -Recurse).Count -ne $web.files.Count + 1) { throw 'Unlisted Web resources exist.' }
    $deps = Get-Content -LiteralPath (Join-Path $package 'QuickExcel.Avalonia.deps.json') -Raw
    if ($deps -match 'Avalonia.DesignerSupport|QuickExcel.Desktop/|QuickExcel.Grid.Wpf|Avalonia.Xpf') { throw 'Forbidden production dependency remains in deps.json.' }
    foreach ($configuration in @('QuickExcel.Avalonia.runtimeconfig.json')) {
        $runtime = Get-Content -LiteralPath (Join-Path $package $configuration) -Raw | ConvertFrom-Json
        if ($runtime.runtimeOptions.PSObject.Properties.Name -contains 'framework') { throw "Payload still needs an installed .NET runtime: $configuration" }
    }
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'README-avalonia-package.txt') -Destination (Join-Path $package 'README.txt')
    $metadata=Join-Path $package 'metadata'; New-Item -ItemType Directory -Path $metadata | Out-Null
    Copy-Item -LiteralPath $lockDirectory -Destination $metadata -Recurse
    Copy-Item -LiteralPath $workerLockDirectory -Destination $metadata -Recurse
    Copy-Item -LiteralPath (Join-Path $root 'QuickExcel.Web/package-lock.json') -Destination (Join-Path $metadata 'web-package-lock.json')
    if ($CodeSigningCertificateThumbprint) {
        $code = @(Get-ChildItem -LiteralPath $package -Recurse -File | Where-Object { $_.Name -match '^QuickExcel\..*\.(exe|dll)$|^quickxl\.exe$' } | ForEach-Object FullName)
        & "$PSScriptRoot/sign-windows-code.ps1" -CertificateThumbprint $CodeSigningCertificateThumbprint -Path $code
    }
    $selfCheck = 'not-run'
    if ($RunSmoke -or $CreateArchive -or $Activate) {
        $checkDirectory = Join-Path $evidence 'self-check'
        $process = Start-Process -FilePath (Join-Path $package 'quickxl.exe') -ArgumentList @('--self-check', ('"{0}"' -f $checkDirectory)) -WorkingDirectory $package -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $evidence 'stdout.log') -RedirectStandardError (Join-Path $evidence 'stderr.log')
        try {
            if (-not $process.WaitForExit(90000)) { throw 'Published self-check exceeded 90 seconds.' }
            $process.WaitForExit()
            $result = Get-Content -LiteralPath (Join-Path $checkDirectory 'result.json') -Raw | ConvertFrom-Json
            if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "Published self-check failed: $($result.error)" }
            $selfCheck = $result
        } finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
    }
    $files = @(Get-ChildItem -LiteralPath $package -File -Recurse | Sort-Object FullName | ForEach-Object { [ordered]@{ path = [IO.Path]::GetRelativePath($package,$_.FullName).Replace('\','/'); size=$_.Length; sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
    $locks = @(& git ls-files '*packages.lock.json' 'QuickExcel.Web/package-lock.json' | ForEach-Object { [ordered]@{ path=$_; sha256=(Get-FileHash -LiteralPath (Join-Path $root $_) -Algorithm SHA256).Hash } })
    [ordered]@{ schemaVersion=1; product='QuickExcel Avalonia'; version=$version; rid='win-x64'; selfContained=$true; diagnostics=$false; scriptEngine=$ScriptEngine; workerDeployment='NativeAOT'; desktopTrimMode='partial'; openIssues=@('JSC-MEM-001: 64 MiB engine hard limit deferred by user; not a progression gate'); sourceHead=(& git rev-parse HEAD); workingTreeDirty=[bool](& git status --porcelain); sdk=(& dotnet --version); webPackages=$web.packages; webManifestSha256=(Get-FileHash (Join-Path $package 'Web/manifest.json') -Algorithm SHA256).Hash; locks=$locks; checks=$selfCheck; files=$files } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $package 'release.json') -Encoding utf8
    if ($CreateArchive) {
        $archive=ArtifactPath "artifacts/avalonia/releases/QuickExcel-Avalonia-$version-win-x64-$runId.zip"
        Compress-Archive -LiteralPath $package -DestinationPath $archive
        (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash + '  ' + [IO.Path]::GetFileName($archive) | Set-Content -LiteralPath ($archive+'.sha256') -Encoding ascii
    }
    if ($Activate) {
        $pointer=ArtifactPath 'artifacts/avalonia/active-windows.json'; $temporary=$pointer+'.'+[Guid]::NewGuid().ToString('N')+'.tmp'
        [ordered]@{ schemaVersion=1; directory=$package; manifestSha256=(Get-FileHash (Join-Path $package 'release.json') -Algorithm SHA256).Hash } | ConvertTo-Json | Set-Content -LiteralPath $temporary -Encoding utf8
        Move-Item -LiteralPath $temporary -Destination $pointer -Force
    }
    Write-Host "PASS: production package $package; evidence $evidence"
} finally { Pop-Location }
} finally {
    # ClickOnce's pinned Mage targets .NET 8 and must keep using the caller's
    # runtime installation after this isolated .NET 10 SDK build has finished.
    $env:PATH = $previousBuildPath
    $env:DOTNET_ROOT = $previousDotnetRoot
    $env:OS = $previousBuildOs
}
