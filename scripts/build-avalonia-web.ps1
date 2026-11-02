[CmdletBinding()]
param([switch]$Probe, [switch]$Tests)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path $PSScriptRoot -Parent
$source = Join-Path $root 'QuickExcel.Web'
$npmRoot = Join-Path $root 'artifacts/avalonia/npm'
New-Item -ItemType Directory -Force $npmRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $source 'package.json'), (Join-Path $source 'package-lock.json') -Destination $npmRoot
Push-Location $npmRoot
try {
    & npm.cmd ci --cache (Join-Path $root 'artifacts/avalonia/npm-cache') --ignore-scripts --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) { throw 'Locked web dependency installation failed.' }
}
finally { Pop-Location }
$buildArguments = @(); if ($Probe) { $buildArguments += '--probe' }; if ($Tests) { $buildArguments += '--tests' }
& node (Join-Path $source 'build.mjs') @buildArguments
if ($LASTEXITCODE -ne 0) { throw 'Offline web build failed.' }
