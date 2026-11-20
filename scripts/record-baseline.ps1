param([string]$EvidenceDirectory = 'artifacts\baseline')

# Capture only: this script does not build, test, publish, or execute workbooks.
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifactsRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot 'artifacts'))
$evidenceRoot = if ([IO.Path]::IsPathRooted($EvidenceDirectory)) {
    [IO.Path]::GetFullPath($EvidenceDirectory)
} else {
    [IO.Path]::GetFullPath((Join-Path $projectRoot $EvidenceDirectory))
}
if (-not $evidenceRoot.StartsWith($artifactsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'EvidenceDirectory must be a directory below this repository artifacts directory.'
}
$captureStartedUtc = [DateTime]::UtcNow.ToString('O')
$captureRoot = Join-Path $evidenceRoot ('baseline-capture-' + [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ') + '-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $captureRoot -Force | Out-Null

function Invoke-CapturedCommand([string]$Executable, [string[]]$Arguments) {
    try {
        $commandOutput = & $Executable @Arguments 2>&1 | ForEach-Object { "$_" }
        return [ordered]@{ command = $Executable; arguments = $Arguments; exitCode = $LASTEXITCODE; output = @($commandOutput) }
    } catch {
        return [ordered]@{ command = $Executable; arguments = $Arguments; exitCode = $null; error = $_.Exception.Message }
    }
}

function Get-FileRecord([string]$Path, [string]$RelativeTo) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [ordered]@{ path = $Path; exists = $false }
    }
    $baselineFile = Get-Item -LiteralPath $Path
    $displayPath = $baselineFile.FullName
    if ($displayPath.StartsWith($RelativeTo + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        $displayPath = $displayPath.Substring($RelativeTo.Length + 1).Replace('\', '/')
    }
    return [ordered]@{
        path = $displayPath
        exists = $true
        length = $baselineFile.Length
        lastWriteTimeUtc = $baselineFile.LastWriteTimeUtc.ToString('O')
        sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    }
}

function Write-Json([string]$Name, $Value) {
    ConvertTo-Json -InputObject $Value -Depth 12 | Set-Content -LiteralPath (Join-Path $captureRoot $Name) -Encoding utf8
}

Push-Location $projectRoot
try {
    $sourceHead = Invoke-CapturedCommand 'git' @('rev-parse', 'HEAD')
    $sourceStatus = Invoke-CapturedCommand 'git' @('status', '--short')
    $sourceNames = @(& git -c core.quotepath=false ls-files --cached --others --exclude-standard)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot enumerate QuickExcel source files.' }
    $sourceFiles = @($sourceNames | Sort-Object -Unique | ForEach-Object { Get-FileRecord (Join-Path $projectRoot $_) $projectRoot })
    Write-Json 'source-files.json' $sourceFiles

    # HEAD plus this patch and copied untracked files identify the local source state.
    $patchPath = Join-Path $captureRoot 'working-tree.patch'
    & git diff HEAD --binary --no-ext-diff "--output=$patchPath"
    if ($LASTEXITCODE -ne 0) { throw 'Cannot capture QuickExcel working-tree diff.' }
    $untrackedNames = @(& git -c core.quotepath=false ls-files --others --exclude-standard)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot enumerate untracked QuickExcel sources.' }
    foreach ($sourceName in $untrackedNames) {
        $sourcePath = [IO.Path]::GetFullPath((Join-Path $projectRoot $sourceName))
        if (-not $sourcePath.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Source path is outside the QuickExcel repository.'
        }
        if (Test-Path -LiteralPath $sourcePath -PathType Leaf) {
            $snapshotPath = Join-Path (Join-Path $captureRoot 'untracked-sources') $sourceName
            New-Item -ItemType Directory -Path (Split-Path $snapshotPath -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $sourcePath -Destination $snapshotPath
        }
    }

    $vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'

    $nativeFiles = @(
        'artifacts/jsc/downloads/webkitgtk-2.54.2.tar.xz',
        'artifacts/jsc/build/win-x64/bin/JavaScriptCore.dll',
        'artifacts/jsc/build/win-x64/identity.json',
        'artifacts/jsc/compiler/identity.txt',
        'QuickExcel.Scripting/native/jsc/bridge.cpp',
        'QuickExcel.Scripting/native/jsc/SOURCES.md',
        'tools/QuickExcel.JscProbe/sources.json',
        'tools/QuickExcel.JscProbe/source-adjustments.json'
    ) | ForEach-Object { Get-FileRecord (Join-Path $projectRoot $_) $projectRoot }
    $evidenceFiles = @(Get-ChildItem -LiteralPath $evidenceRoot -File -Recurse | Where-Object {
        $_.FullName.Substring($evidenceRoot.Length + 1) -notmatch '(^|[\\/])baseline-capture-[^\\/]+([\\/]|$)'
    } | Sort-Object FullName | ForEach-Object { Get-FileRecord $_.FullName $evidenceRoot })
    Write-Json 'evidence-files.json' $evidenceFiles

    Write-Json 'metadata.json' ([ordered]@{
        schemaVersion = 1
        captureStartedUtc = $captureStartedUtc
        captureFinishedUtc = [DateTime]::UtcNow.ToString('O')
        interpretation = 'Observed files and environment only. This capture does not execute or attest builds/tests. Use run logs, timestamps and process exit codes to establish fresh validation.'
        source = [ordered]@{ repository = $projectRoot; head = $sourceHead; status = $sourceStatus; manifest = 'source-files.json'; diff = 'working-tree.patch'; untrackedSources = 'untracked-sources' }
        environment = [ordered]@{
            os = [Environment]::OSVersion.VersionString
            is64BitOperatingSystem = [Environment]::Is64BitOperatingSystem
            is64BitProcess = [Environment]::Is64BitProcess
            powershell = $PSVersionTable.PSVersion.ToString()
            dotnetSdk = Invoke-CapturedCommand 'dotnet' @('--version')
            dotnetRuntimes = Invoke-CapturedCommand 'dotnet' @('--list-runtimes')
            dotnetInfoLimitation = 'dotnet --info is intentionally not invoked: the 2026-10-22 sandbox read threw InstallerBase TypeInitializationException while inspecting workloads. SDK version and installed runtimes are captured separately.'
            cmake = Invoke-CapturedCommand 'cmake' @('--version')
            visualStudio = Invoke-CapturedCommand $vswherePath @('-latest', '-products', '*', '-requires', 'Microsoft.VisualStudio.Component.VC.Tools.x86.x64', '-property', 'installationVersion')
        }
        native = [ordered]@{
            engine = 'JavaScriptCore/2.54.2/b95c91b1/qx-async-3'
            expectedArchiveSha256 = 'A76345B391BEF6211FB5AB6CFAB2E86B285D18CDA17CBB9ADE595F1252B86065'
            interpretation = 'Existing binaries are hashed, not rebuilt by this script. Matching archive hashes do not prove that an observed DLL was rebuilt during this validation run.'
            buildContract = 'Pinned JSCOnly/ICU source build; JIT/DFG/mimalloc; no FTL/Wasm/inspector. Full compiler, options and binary hashes are in native identity. JSC-MEM-001 open.'
            files = @($nativeFiles)
        }
        publish = [ordered]@{ configuration = 'Release'; architecture = 'x64'; selfContained = $true; requiredRuntime = 'WebView2 Evergreen (Windows); .NET included' }
        evidence = [ordered]@{ directory = $evidenceRoot; manifest = 'evidence-files.json'; count = $evidenceFiles.Count }
    })
    Write-Host "Baseline capture: $captureRoot"
} finally {
    Pop-Location
}
