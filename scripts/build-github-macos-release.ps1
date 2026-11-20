[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (-not $IsMacOS -or [Runtime.InteropServices.RuntimeInformation]::OSArchitecture -ne 'Arm64') { throw 'Run the macOS release build on macOS arm64.' }
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$output = Join-Path $root 'artifacts/github-release'
$evidence = Join-Path $root 'artifacts/github-evidence/macos'
$releases = Join-Path $root 'artifacts/avalonia/platform-releases'
$version = ([xml](Get-Content -LiteralPath (Join-Path $root 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj') -Raw)).SelectSingleNode('/Project/PropertyGroup/Version').InnerText
$archive = Join-Path $output "QuickExcel-$version-osx-arm64.tar.gz"
if ((Test-Path -LiteralPath $archive) -or (Test-Path -LiteralPath (Join-Path $evidence 'self-check'))) { throw 'Use a fresh checkout/output directory for a release build.' }
New-Item -ItemType Directory -Path $output,$evidence,$releases -Force | Out-Null
$before = @(Get-ChildItem -LiteralPath $releases -Directory | ForEach-Object FullName)
Push-Location $root
try {
    bash tools/QuickExcel.JscProbe/build-macos.sh
    if ($LASTEXITCODE -ne 0) { throw 'macOS native engine build failed.' }
    bash QuickExcel.ScriptWorker/native/build-macos.sh
    if ($LASTEXITCODE -ne 0) { throw 'macOS sandbox build failed.' }
    & "$PSScriptRoot/publish-avalonia-unix.ps1" -Rid osx-arm64
    if (-not $?) { throw 'macOS production publish failed.' }
    $created = @(Get-ChildItem -LiteralPath $releases -Directory | Where-Object { $_.Name.StartsWith('osx-arm64-') -and $_.FullName -notin $before })
    if ($created.Count -ne 1) { throw 'Cannot identify the fresh macOS release directory.' }
    $package = Join-Path $created[0].FullName 'application'
    $manifestPath = Join-Path $package 'release.json'
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if ($manifest.version -ne $version -or $manifest.rid -ne 'osx-arm64' -or $manifest.diagnostics) { throw 'Unexpected macOS release identity.' }
    $checkDirectory = Join-Path $evidence 'self-check'
    $process = Start-Process -FilePath (Join-Path $package 'QuickExcel.Avalonia') -ArgumentList @('--self-check', ('"{0}"' -f $checkDirectory)) -WorkingDirectory $package -PassThru -RedirectStandardOutput (Join-Path $evidence 'stdout.log') -RedirectStandardError (Join-Path $evidence 'stderr.log')
    try {
        if (-not $process.WaitForExit(180000)) { throw 'Published macOS self-check exceeded 180 seconds.' }
        $process.WaitForExit()
        $result = Get-Content -LiteralPath (Join-Path $checkDirectory 'result.json') -Raw | ConvertFrom-Json
        if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "Published macOS self-check failed: $($result.error)" }
    } finally {
        if (-not $process.HasExited) { $process.Kill($true); [void]$process.WaitForExit(5000) }
        $process.Dispose()
    }
    $manifest | Add-Member -NotePropertyName checks -NotePropertyValue $result -Force
    $manifest.acceptance = 'Production self-check passed on the build runner; signing, notarization and target-device acceptance remain separate.'
    $manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding utf8
    Copy-Item -LiteralPath $manifestPath -Destination $evidence

    dotnet test QuickExcel.Core.Test/QuickExcel.Core.Test.csproj -c Release -r osx-arm64 -p:SelfContained=true -p:QuickExcelNativeRid=osx-arm64 "-p:QuickExcelPublishedWorkerDirectory=$package/ScriptWorker" --artifacts-path (Join-Path $root 'artifacts/github-test-build/macos') --filter 'FullyQualifiedName~MacScriptWorkerTests|FullyQualifiedName~M4PolicyTests|FullyQualifiedName~ScriptExecutionPolicyTests|FullyQualifiedName~ScriptSigningPreferencesTests' --logger 'trx;LogFileName=release-contracts.trx' --results-directory $evidence -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'macOS published-worker contract checks failed.' }
    python3 "$PSScriptRoot/archive-unix-package.py" $package $archive
    if ($LASTEXITCODE -ne 0) { throw 'macOS archive verification failed.' }
    Write-Host "PASS: $archive"
} finally { Pop-Location }
