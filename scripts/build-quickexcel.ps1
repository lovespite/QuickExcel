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
        dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c $Configuration --artifacts-path $artifactsPath -v minimal
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

    foreach ($required in @('QuickExcel.exe', 'QuickExcel.dll', 'QuickExcel.runtimeconfig.json', 'quickexcel-js.dll', 'QUICKJS-LICENSE.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $publishPath $required) -PathType Leaf)) { throw "Published application is missing $required" }
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
            foreach ($resultName in @('storage-smoke-result.txt', 'close-smoke-result.txt', 'calculation-smoke-result.txt', 'date-display-smoke-result.txt')) {
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
    $dependencyRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot '..\ExcelIO'))
    [xml]$excelIoProject = Get-Content -LiteralPath (Join-Path $dependencyRoot 'ExcelIO\ExcelIO.csproj') -Raw
    [xml]$formulaProject = Get-Content -LiteralPath (Join-Path $dependencyRoot 'ExcelIO.Formula\ExcelIO.Formula.csproj') -Raw
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
        excelIoSource = Read-GitState $dependencyRoot
        sdkVersion = $sdkVersion
        installedRuntimesOnBuildMachine = $runtimes
        nativeRuntime = [ordered]@{
            name = 'QuickJS-NG'
            version = $nativeVersionMatch.Groups[1].Value
            sourceArchiveSha256 = $nativeHashMatch.Groups[1].Value.ToUpperInvariant()
            publishedDllSha256 = (Get-FileHash -LiteralPath (Join-Path $publishPath 'quickexcel-js.dll') -Algorithm SHA256).Hash
        }
        verification = [ordered]@{ coreTests = $unitTests; nativeSmoke = $nativeSmoke; wpfSmoke = $desktopSmoke; wpfResult = $smokeResult }
    }
    $metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $publishPath 'BUILD-METADATA.json') -Encoding UTF8
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\RELEASE-NOTES.md') -Destination (Join-Path $publishPath 'RELEASE-NOTES.md')
    $supportDocsPath = Join-Path $publishPath 'docs'
    New-Item -ItemType Directory -Force -Path $supportDocsPath | Out-Null
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\09-m2-files-and-types.md') -Destination (Join-Path $supportDocsPath '09-m2-files-and-types.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Desktop\docs\10-m3-calculation.md') -Destination (Join-Path $supportDocsPath '10-m3-calculation.md')
    Copy-Item -LiteralPath (Join-Path $projectRoot 'QuickExcel.Scripting\native\SOURCES.md') -Destination (Join-Path $publishPath 'QUICKJS-SOURCES.md')
    $notices = @"
QuickExcel third-party component provenance

QuickJS-NG $($nativeVersionMatch.Groups[1].Value)
Source: https://github.com/quickjs-ng/quickjs
Declared license: MIT. Complete copyright and license text: QUICKJS-LICENSE.txt.
Source archive pin and build procedure: QUICKJS-SOURCES.md.

ExcelIO $($excelIoProject.Project.PropertyGroup.Version)
Author recorded by project: $($excelIoProject.Project.PropertyGroup.Authors)
Repository recorded by project: $($excelIoProject.Project.PropertyGroup.RepositoryUrl)
PackageLicenseExpression recorded by project: $($excelIoProject.Project.PropertyGroup.PackageLicenseExpression)

ExcelIO.Formula $($formulaProject.Project.PropertyGroup.Version)
Author recorded by project: $($formulaProject.Project.PropertyGroup.Authors)
Repository recorded by project: $($formulaProject.Project.PropertyGroup.RepositoryUrl)
PackageLicenseExpression recorded by project: $($formulaProject.Project.PropertyGroup.PackageLicenseExpression)

The ExcelIO .NET project files declare the licenses above. No standalone .NET license text was supplied in the dependency source tree; no copyright year or license text is inferred from its separate npm package. This application package includes compiled dependency assemblies, not sibling repository sources.
"@
    Set-Content -LiteralPath (Join-Path $publishPath 'THIRD-PARTY-NOTICES.txt') -Value $notices -Encoding UTF8

    # Copy only this fresh publish tree; never enumerate old output files into a release package.
    New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
    foreach ($entry in Get-ChildItem -LiteralPath $publishPath -Force) {
        Copy-Item -LiteralPath $entry.FullName -Destination $outputPath -Recurse -Force
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
