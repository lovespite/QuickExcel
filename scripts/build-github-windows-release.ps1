[CmdletBinding()]
param([ValidateRange(1, 16)][int]$Parallel = 2)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (-not $IsWindows) { throw 'Run the Windows release build on Windows x64.' }
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$package = Join-Path $root 'artifacts/github-package/windows/application'
$output = Join-Path $root 'artifacts/github-release'
$evidence = Join-Path $root 'artifacts/github-evidence/windows'
$version = ([xml](Get-Content -LiteralPath (Join-Path $root 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj') -Raw)).SelectSingleNode('/Project/PropertyGroup/Version').InnerText
$archive = Join-Path $output "QuickExcel-$version-win-x64.zip"
if ((Test-Path -LiteralPath $package) -or (Test-Path -LiteralPath $archive)) { throw 'Use a fresh checkout/output directory for a release build.' }
New-Item -ItemType Directory -Path $output,$evidence -Force | Out-Null
Push-Location $root
try {
    # Run in a child shell: Visual Studio's native build environment must not
    # replace setup-dotnet's SDK selection for the managed publish below.
    & pwsh -NoProfile -File tools/QuickExcel.JscProbe/build.ps1 -Parallel $Parallel
    if ($LASTEXITCODE -ne 0) { throw 'Windows native build failed.' }
    & "$PSScriptRoot/build-avalonia-release.ps1" -SkipTests -RunSmoke -OutputDirectory $package
    if (-not $?) { throw 'Windows production publish/self-check failed.' }
    $manifest = Get-Content -LiteralPath (Join-Path $package 'release.json') -Raw | ConvertFrom-Json
    if ($manifest.version -ne $version -or $manifest.rid -ne 'win-x64' -or -not $manifest.checks.passed) { throw 'Unverified Windows release package.' }
    Copy-Item -LiteralPath (Join-Path $package 'release.json') -Destination $evidence

    # Exercise the worker that will ship, including its Native AOT deployment.
    # The package's production self-check also covers storage, Web and rendering.
    $filter = 'FullyQualifiedName~M4PolicyTests|FullyQualifiedName~ScriptExecutionPolicyTests|FullyQualifiedName~ScriptSigningPreferencesTests|FullyQualifiedName~M4WorkerTests.WorkerHasSeparateProcessAndPreservesScalarTypesAndBatchOrder|FullyQualifiedName~M4WorkerTests.ContextsAndLogsAreIsolatedAndConservativeRoundsNeverReuseUdfResults'
    dotnet test QuickExcel.Core.Test/QuickExcel.Core.Test.csproj -c Release --artifacts-path (Join-Path $root 'artifacts/github-test-build/windows') "-p:QuickExcelPublishedWorkerDirectory=$package/ScriptWorker" --filter $filter --logger 'trx;LogFileName=release-contracts.trx' --results-directory $evidence -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'Windows published-worker contract checks failed.' }

    Compress-Archive -LiteralPath $package -DestinationPath $archive
    (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + [IO.Path]::GetFileName($archive) | Set-Content -LiteralPath ($archive + '.sha256') -Encoding ascii
    Write-Host "PASS: $archive"
} finally { Pop-Location }
