[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Release',
    [switch]$SkipTests,
    [string]$OutputDirectory = 'artifacts\QuickExcel',
    [switch]$RunSmoke,
    [switch]$CreateArchive
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$projectRoot = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifactsRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot 'artifacts'))
$artifactsPrefix = $artifactsRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
$artifactsPath = Join-Path $artifactsRoot 'build'

function Resolve-ArtifactDirectory([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { throw 'OutputDirectory cannot be blank.' }
    $candidate = if ([IO.Path]::IsPathRooted($Value)) { $Value } else { Join-Path $projectRoot $Value }
    $full = [IO.Path]::GetFullPath($candidate)
    if (-not $full.StartsWith($artifactsPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'OutputDirectory must be a child directory inside this repository artifacts directory.'
    }
    # Existing junctions or symbolic links must not redirect generated writes outside the repository.
    $check = $full
    while ($check -and $check.Length -ge $artifactsRoot.Length) {
        if (Test-Path -LiteralPath $check) {
            $item = Get-Item -LiteralPath $check -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Generated output cannot pass through a reparse point: $check"
            }
            if (-not $item.PSIsContainer) { throw "Output path is not a directory: $check" }
        }
        $check = Split-Path $check -Parent
    }
    return $full
}

function Read-GitState([string]$Root) {
    $head = (& git -c "safe.directory=$Root" -C $Root rev-parse HEAD 2>$null | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { throw "Cannot read Git HEAD in $Root" }
    $status = (& git -c "safe.directory=$Root" -C $Root status --porcelain 2>$null | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { throw "Cannot read Git working-tree state in $Root" }
    return [ordered]@{ head = $head; workingTreeDirty = -not [string]::IsNullOrWhiteSpace($status) }
}

$outputPath = Resolve-ArtifactDirectory $OutputDirectory
$runId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$publishPath = Resolve-ArtifactDirectory (Join-Path $artifactsRoot "publish\$runId")
$smokePath = Resolve-ArtifactDirectory (Join-Path $artifactsRoot "smoke\$runId")
$testResultsPath = Resolve-ArtifactDirectory (Join-Path $artifactsRoot "test-results\$runId")
$releasePath = Resolve-ArtifactDirectory (Join-Path $artifactsRoot 'releases')
$desktopProject = Join-Path $projectRoot 'QuickExcel.Desktop\QuickExcel.Desktop.csproj'
[xml]$desktopProjectXml = Get-Content -LiteralPath $desktopProject -Raw
$version = [string]$desktopProjectXml.Project.PropertyGroup.Version
if ($version -notmatch '^\d+\.\d+\.\d+(?:[-+][A-Za-z0-9.-]+)?$') { throw 'A valid release Version is required in the desktop project.' }
$shouldRunSmoke = $RunSmoke -or $CreateArchive
$unitTests = 'skipped'
$nativeSmoke = 'skipped'
$desktopSmoke = 'skipped'
$smokeResult = $null

Push-Location $projectRoot
try {
    New-Item -ItemType Directory -Force -Path $publishPath | Out-Null
    if (-not $SkipTests) {
        dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c $Configuration --artifacts-path $artifactsPath -v minimal --logger 'trx;LogFileName=core-tests.trx' --results-directory $testResultsPath
        if ($LASTEXITCODE -ne 0) { throw 'QuickExcel tests failed.' }
        $unitTests = 'passed'
    }
    if (-not $SkipTests -or $shouldRunSmoke) {
        dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c $Configuration --artifacts-path $artifactsPath
        if ($LASTEXITCODE -ne 0) { throw 'QuickExcel native runtime smoke failed.' }
        $nativeSmoke = 'passed'
    }
    dotnet publish .\QuickExcel.Desktop\QuickExcel.Desktop.csproj -c $Configuration --artifacts-path $artifactsPath -r win-x64 --self-contained false -o $publishPath -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'QuickExcel publish failed.' }

    foreach ($required in @('QuickExcel.exe', 'QuickExcel.dll', 'QuickExcel.runtimeconfig.json', 'quickexcel-js.dll', 'QUICKJS-LICENSE.txt',
        'QuickExcel.Grid.Wpf.dll', 'FASTWPFGRID-LICENSE.txt', 'WRITEABLEBITMAPEX-LICENSE.txt',
        'QuickExcel.Rendering.dll', 'QuickExcel.Rendering.Skia.dll', 'QuickExcel.RichTextKit.dll', 'SkiaSharp.dll',
        'libSkiaSharp.dll', 'HarfBuzzSharp.dll', 'libHarfBuzzSharp.dll', 'BitMiracle.LibTiff.NET.dll', 'UglyToad.PdfPig.dll',
        'RICH-TEXT-KIT-LICENSE.txt', 'SKIASHARP-LICENSE.txt', 'HARFBUZZSHARP-LICENSE.txt',
        'SKIA-THIRD-PARTY-NOTICES.txt', 'HARFBUZZ-THIRD-PARTY-NOTICES.txt', 'LIBTIFF-LICENSE.txt', 'PDFPIG-LICENSE.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $publishPath $required) -PathType Leaf)) { throw "Published application is missing $required" }
    }

    foreach ($required in @('QuickExcel.ScriptWorker.exe', 'QuickExcel.ScriptWorker.dll', 'QuickExcel.ScriptWorker.runtimeconfig.json',
        'QuickExcel.ScriptWorker.deps.json', 'QuickExcel.ScriptContracts.dll', 'QuickExcel.Scripting.dll', 'quickexcel-js.dll', 'QUICKJS-LICENSE.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $publishPath "ScriptWorker\$required") -PathType Leaf)) { throw "Published script worker is missing $required" }
    }
    if (Test-Path -LiteralPath (Join-Path $publishPath 'WorkerFixture')) { throw 'Test worker fixture must not be published.' }
    if (Get-ChildItem -LiteralPath (Join-Path $publishPath 'ScriptWorker') -Recurse -File | Where-Object {
        $_.Name -match 'Skia|HarfBuzz|RichTextKit|LibTiff|PdfPig|Rendering' }) { throw 'Rendering dependencies must not enter the script worker payload.' }
    if (Get-ChildItem -LiteralPath $publishPath -Recurse -File | Where-Object { $_.Name -like '*HostFixture*' -or $_.Name -like 'WorkerFixture*' }) {
        throw 'Test fixtures must not be published.'
    }
    if (Get-ChildItem -LiteralPath $publishPath -Recurse -File | Where-Object { $_.Name -like 'QuickExcel.SandboxProbe*' }) {
        throw 'Development sandbox probe must not be published.'
    }
    if (Get-ChildItem -LiteralPath $publishPath -Recurse -File | Where-Object { $_.Name -like 'ExcelIO*' }) {
        throw 'Retired ExcelIO dependencies must not be published.'
    }

    if ($shouldRunSmoke) {
        New-Item -ItemType Directory -Force -Path $smokePath | Out-Null
        $desktopProcess = $null
        try {
            $desktopProcess = Start-Process -FilePath (Join-Path $publishPath 'QuickExcel.exe') -ArgumentList @('--smoke-test', ('"{0}"' -f $smokePath)) -WorkingDirectory $publishPath -WindowStyle Hidden -PassThru
            if (-not $desktopProcess.WaitForExit(45000)) {
                if (-not $desktopProcess.HasExited) { $desktopProcess.Kill(); [void]$desktopProcess.WaitForExit(5000) }
                throw "QuickExcel WPF smoke exceeded its 45 second deadline. Evidence: $smokePath"
            }
            $desktopProcess.WaitForExit()
            $desktopProcess.Refresh()
            if ($desktopProcess.ExitCode -ne 0) { throw "QuickExcel WPF smoke failed with exit code $($desktopProcess.ExitCode). Evidence: $smokePath" }
            $resultPath = Join-Path $smokePath 'smoke-result.txt'
            if (-not (Test-Path -LiteralPath $resultPath -PathType Leaf)) { throw "QuickExcel WPF smoke exited without its result file: $smokePath" }
            $smokeResult = (Get-Content -LiteralPath $resultPath -Raw).Trim()
            if (-not $smokeResult.StartsWith('PASS:', [StringComparison]::Ordinal)) { throw "QuickExcel WPF smoke did not report PASS: $smokePath" }
            foreach ($resultName in @('storage-smoke-result.txt', 'close-smoke-result.txt', 'calculation-smoke-result.txt', 'date-display-smoke-result.txt', 'worker-smoke-result.txt', 'actions-smoke-result.txt', 'interaction-smoke-result.txt', 'workspace-smoke-result.txt', 'product-smoke-result.txt')) {
                $resultFile = Join-Path $smokePath $resultName
                if (-not (Test-Path -LiteralPath $resultFile -PathType Leaf) -or
                    -not (Get-Content -LiteralPath $resultFile -Raw).StartsWith('PASS:', [StringComparison]::Ordinal)) {
                    throw "Missing fresh smoke result: $resultName"
                }
            }
            $smokeResult += ' ' + (Get-Content -LiteralPath (Join-Path $smokePath 'calculation-smoke-result.txt') -Raw).Trim()
            $smokeResult += ' ' + (Get-Content -LiteralPath (Join-Path $smokePath 'date-display-smoke-result.txt') -Raw).Trim()
            $desktopSmoke = 'passed'
        }
        finally {
            if ($null -ne $desktopProcess) {
                if (-not $desktopProcess.HasExited) { $desktopProcess.Kill(); [void]$desktopProcess.WaitForExit(5000) }
                $desktopProcess.Dispose()
            }
        }
    }

    $sdkVersion = (& dotnet --version | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Could not read the .NET SDK version.' }
    $runtimes = @(& dotnet --list-runtimes)
    if ($LASTEXITCODE -ne 0) { throw 'Could not read installed .NET runtimes.' }
    $nativeBuildText = Get-Content -LiteralPath (Join-Path $projectRoot 'QuickExcel.Scripting\native\build.ps1') -Raw
    $nativeVersionMatch = [regex]::Match($nativeBuildText, '\$version\s*=\s*''([^'']+)''')
    $nativeHashMatch = [regex]::Match($nativeBuildText, '\$expectedHash\s*=\s*''([0-9A-Fa-f]{64})''')
    if (-not $nativeVersionMatch.Success -or -not $nativeHashMatch.Success) { throw 'Cannot read the pinned native runtime source version/hash.' }
    $metadata = [ordered]@{
        product = 'QuickExcel'
        version = $version
        buildUtc = [DateTime]::UtcNow.ToString('o')
        configuration = $Configuration
        runtimeIdentifier = 'win-x64'
        deployment = 'framework-dependent'
        selfContained = $false
        requiredRuntime = '.NET 10 Windows Desktop Runtime (x64)'
        cleanMachineValidated = $false
        source = Read-GitState $projectRoot
        gridSource = [ordered]@{
            upstreamCommit = (Get-Content -LiteralPath (Join-Path $projectRoot 'QuickExcel.Grid.Wpf\UPSTREAM-SOURCES.json') -Raw | ConvertFrom-Json).commit
            publishedDllSha256 = (Get-FileHash -LiteralPath (Join-Path $publishPath 'QuickExcel.Grid.Wpf.dll') -Algorithm SHA256).Hash
            defaultSurface = 'FastGridSurface'
        }
        rendering = [ordered]@{
            backend = 'SkiaSharp CPU'
            pdf = 'Direct Skia text/vector PDF'
            packages = @((Get-Content -LiteralPath (Join-Path $projectRoot 'QuickExcel.Rendering.Skia\packages.lock.json') -Raw | ConvertFrom-Json).dependencies.'net10.0'.PSObject.Properties |
                Where-Object { $_.Value.type -in @('Direct', 'Transitive') } | ForEach-Object {
                    [ordered]@{ name = $_.Name; version = $_.Value.resolved; contentHash = $_.Value.contentHash }
                })
            nativeLibraries = @('libSkiaSharp.dll', 'libHarfBuzzSharp.dll') | ForEach-Object {
                [ordered]@{ name = $_; sha256 = (Get-FileHash -LiteralPath (Join-Path $publishPath $_) -Algorithm SHA256).Hash }
            }
            richTextKitSource = '0.6.170 / 6362fcdc33ea43b359ada2da946bf025a6d14630 (local adaptations; see third_party/RichTextKit/SOURCES.md)'
        }
        sdkVersion = $sdkVersion
        installedRuntimesOnBuildMachine = $runtimes
        nativeRuntime = [ordered]@{
            name = 'QuickJS-NG'
            version = $nativeVersionMatch.Groups[1].Value
            sourceArchiveSha256 = $nativeHashMatch.Groups[1].Value.ToUpperInvariant()
            publishedDllSha256 = (Get-FileHash -LiteralPath (Join-Path $publishPath 'quickexcel-js.dll') -Algorithm SHA256).Hash
        }
        verification = [ordered]@{ coreTests = $unitTests; coreTestResults = $(if ($unitTests -eq 'passed') { Join-Path $testResultsPath 'core-tests.trx' } else { $null }); nativeSmoke = $nativeSmoke; wpfSmoke = $desktopSmoke; wpfResult = $smokeResult }
        scriptWorker = [ordered]@{ protocolVersion = 4; policy = 'conservative-worker-v1'; maxMessageBytes = 16777216; maxCalculationBatchCalls = 85; isolationMode = 'appcontainer-no-capabilities-v1'; osAccessIsolationValidated = $false }
    }
    $metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $publishPath 'BUILD-METADATA.json') -Encoding UTF8
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\RELEASE-NOTES.md') -Destination (Join-Path $publishPath 'RELEASE-NOTES.md')
    $supportDocsPath = Join-Path $publishPath 'docs'
    New-Item -ItemType Directory -Force -Path $supportDocsPath | Out-Null
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\09-m2-files-and-types.md') -Destination (Join-Path $supportDocsPath '09-m2-files-and-types.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\10-m3-calculation.md') -Destination (Join-Path $supportDocsPath '10-m3-calculation.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\11-m4-worker.md') -Destination (Join-Path $supportDocsPath '11-m4-worker.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\12-m4-modules.md') -Destination (Join-Path $supportDocsPath '12-m4-modules.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\13-m4-actions.md') -Destination (Join-Path $supportDocsPath '13-m4-actions.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\14-m4-batching.md') -Destination (Join-Path $supportDocsPath '14-m4-batching.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\15-m4-os-isolation-adr.md') -Destination (Join-Path $supportDocsPath '15-m4-os-isolation-adr.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\16-m4-recovery-and-deployment.md') -Destination (Join-Path $supportDocsPath '16-m4-recovery-and-deployment.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\17-m5-grid-and-input.md') -Destination (Join-Path $supportDocsPath '17-m5-grid-and-input.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\18-m5-structure-and-workspace.md') -Destination (Join-Path $supportDocsPath '18-m5-structure-and-workspace.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\19-m5-product-workspace.md') -Destination (Join-Path $supportDocsPath '19-m5-product-workspace.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\20-script-signatures.md') -Destination (Join-Path $supportDocsPath '20-script-signatures.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\21-excel-ribbon.md') -Destination (Join-Path $supportDocsPath '21-excel-ribbon.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\22-fastwpfgrid-and-data-operations.md') -Destination (Join-Path $supportDocsPath '22-fastwpfgrid-and-data-operations.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\23-fast-grid-implementation.md') -Destination (Join-Path $supportDocsPath '23-fast-grid-implementation.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\59-skia-rendering-and-pdf-plan.md') -Destination (Join-Path $supportDocsPath '59-skia-rendering-and-pdf-plan.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\60-skia-rendering-and-pdf-implementation.md') -Destination (Join-Path $supportDocsPath '60-skia-rendering-and-pdf-implementation.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Rendering.Skia\SOURCES.md') -Destination (Join-Path $publishPath 'SKIA-RENDERING-SOURCES.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'third_party\RichTextKit\SOURCES.md') -Destination (Join-Path $publishPath 'RICH-TEXT-KIT-SOURCES.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Grid.Wpf\SOURCES.md') -Destination (Join-Path $publishPath 'FASTWPFGRID-SOURCES.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Grid.Wpf\UPSTREAM-SOURCES.json') -Destination (Join-Path $publishPath 'FASTWPFGRID-UPSTREAM-SOURCES.json')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Scripting\native\SOURCES.md') -Destination (Join-Path $publishPath 'QUICKJS-SOURCES.md')
    $notices = @"
QuickExcel third-party component provenance

FastWpfGrid and its embedded WriteableBitmapEx fork
Source: https://github.com/janproch/fastwpfgrid
Pinned commit and local changes: FASTWPFGRID-SOURCES.md and FASTWPFGRID-UPSTREAM-SOURCES.json.
Declared licenses: MIT. Complete texts: FASTWPFGRID-LICENSE.txt and WRITEABLEBITMAPEX-LICENSE.txt.

FluentIcons.Wpf / FluentIcons.Common 2.1.344
Source: https://github.com/davidxuang/FluentIcons
Upstream artwork: https://github.com/microsoft/fluentui-system-icons
Declared license: MIT. Complete copyright and license text: FLUENT-ICONS-LICENSE.txt.

QuickJS-NG $($nativeVersionMatch.Groups[1].Value)
Source: https://github.com/quickjs-ng/quickjs
Declared license: MIT. Complete copyright and license text: QUICKJS-LICENSE.txt.
Source archive pin and build procedure: QUICKJS-SOURCES.md.

Shared rendering: SkiaSharp 4.154.0 / HarfBuzzSharp 14.2.1.400,
RichTextKit 0.6.170 (locally adapted source), LibTiff.NET 2.4.660, PdfPig 0.1.16.
Pinned sources, local modifications and native packaging: SKIA-RENDERING-SOURCES.md,
RICH-TEXT-KIT-SOURCES.md, and BUILD-METADATA.json.
Declared licenses: MIT (SkiaSharp/HarfBuzzSharp bindings), Apache-2.0 (RichTextKit/PdfPig),
BSD-style with upstream notices (LibTiff.NET). Complete texts: SKIASHARP-LICENSE.txt,
HARFBUZZSHARP-LICENSE.txt, RICH-TEXT-KIT-LICENSE.txt, LIBTIFF-LICENSE.txt, PDFPIG-LICENSE.txt.
Native third-party licenses: SKIA-THIRD-PARTY-NOTICES.txt, HARFBUZZ-THIRD-PARTY-NOTICES.txt.
Test fonts and independent PDF verification tools are not included in this package.

"@
    Set-Content -LiteralPath (Join-Path $publishPath 'THIRD-PARTY-NOTICES.txt') -Value $notices -Encoding UTF8

    # Copy only this fresh publish tree; never enumerate old output files into a release package.
    New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
    foreach ($entry in Get-ChildItem -LiteralPath $publishPath -Force) {
        Copy-Item -LiteralPath $entry.FullName -Destination $outputPath -Recurse -Force
    }
    # Older installations may still contain these retired assemblies after the overlay copy.
    foreach ($retiredName in @('ExcelIO.dll', 'ExcelIO.pdb', 'ExcelIO.xml', 'ExcelIO.Formula.dll', 'ExcelIO.Formula.pdb', 'ExcelIO.Formula.xml')) {
        $retiredPath = Join-Path $outputPath $retiredName
        if (Test-Path -LiteralPath $retiredPath -PathType Leaf) { Remove-Item -LiteralPath $retiredPath -Force }
    }
    if ($CreateArchive) {
        New-Item -ItemType Directory -Force -Path $releasePath | Out-Null
        $archiveName = "QuickExcel-$version-win-x64-$runId.zip"
        $archivePath = Join-Path $releasePath $archiveName
        if (Test-Path -LiteralPath $archivePath) { throw "Release archive already exists; refusing to overwrite: $archivePath" }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::CreateFromDirectory($publishPath, $archivePath, [IO.Compression.CompressionLevel]::Optimal, $false)
        $archiveHash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash
        Set-Content -LiteralPath ($archivePath + '.sha256') -Value "$archiveHash  $archiveName" -Encoding ASCII
        Write-Host "Release: $archivePath"
        Write-Host "SHA-256: $archiveHash"
    }
    Write-Host "Ready: $(Join-Path $outputPath 'QuickExcel.exe')"
    Write-Host 'Requires .NET 10 Windows Desktop Runtime (x64); this is a framework-dependent package.'
    if ($shouldRunSmoke) { Write-Host "Smoke evidence: $smokePath" }
}
finally { Pop-Location }
