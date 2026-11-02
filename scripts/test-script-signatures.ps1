[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts = [IO.Path]::GetFullPath((Join-Path $root 'artifacts'))
$package = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($PackageDirectory)) { $PackageDirectory } else { Join-Path $root $PackageDirectory }))
if (-not $package.StartsWith($artifacts + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'PackageDirectory must be under this repository artifacts directory.'
}
# Historical rollback checks only; current UI tests use test-avalonia-*.ps1.
$identity=Get-Content -LiteralPath (Join-Path $package 'release.json') -Raw | ConvertFrom-Json
if ($identity.product -ne 'QuickExcel WPF rollback') { throw 'Choose an explicitly archived WPF rollback package.' }
$exe = Join-Path $package 'QuickExcel.exe'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Missing executable: $exe" }
$run = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$evidence = Join-Path $artifacts "validation\script-signatures\$run"
New-Item -ItemType Directory -Path $evidence | Out-Null
$process = $null
try {
    $process = Start-Process -FilePath $exe -ArgumentList @('--signature-smoke', ('"{0}"' -f $evidence)) -WorkingDirectory $package -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(60000)) { throw "Signature smoke exceeded 60 seconds. Evidence: $evidence" }
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "Signature smoke failed with exit code $($process.ExitCode). Evidence: $evidence" }
    $result = Join-Path $evidence 'signature-smoke-result.txt'
    if (-not (Test-Path -LiteralPath $result) -or -not (Get-Content -LiteralPath $result -Raw).StartsWith('PASS:')) {
        throw "Missing fresh signature result: $evidence"
    }
    Get-Content -LiteralPath $result
    [ordered]@{
        package = $package
        desktopSha256 = (Get-FileHash -LiteralPath (Join-Path $package 'QuickExcel.dll') -Algorithm SHA256).Hash
        infrastructureSha256 = (Get-FileHash -LiteralPath (Join-Path $package 'QuickExcel.Infrastructure.dll') -Algorithm SHA256).Hash
        evidence = $evidence
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'run.json') -Encoding utf8
}
finally {
    if ($null -ne $process) {
        if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
        $process.Dispose()
    }
}
Write-Host "Signature evidence: $evidence"
