[CmdletBinding()]
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$package = (Get-Item -LiteralPath $PackageDirectory).FullName
$worker = Join-Path $package 'ScriptWorker'
$engineFiles = @('JavaScriptCore.dll','icudt77.dll','icuuc77.dll','icuin77.dll','msvcp140.dll','msvcp140_2.dll','vcruntime140.dll','vcruntime140_1.dll','typescript.bundle.json','JSC-BUILD-IDENTITY.json')
foreach ($name in $engineFiles) {
    if (Test-Path -LiteralPath (Join-Path $package $name)) { throw "Duplicated worker-only resource: $name" }
    if (-not (Test-Path -LiteralPath (Join-Path $worker $name) -PathType Leaf)) { throw "Missing worker resource: $name" }
}
foreach ($name in @('coreclr.dll','clrjit.dll','hostfxr.dll','hostpolicy.dll','System.Private.CoreLib.dll','QuickExcel.ScriptWorker.dll','QuickExcel.ScriptWorker.deps.json','QuickExcel.ScriptWorker.runtimeconfig.json')) {
    if (Test-Path -LiteralPath (Join-Path $worker $name)) { throw "Managed runtime entered the Native AOT worker: $name" }
}
# Check the native PE architecture/header as well as the absent CLR payload.
# The required production self-check proves that this executable runs in the
# isolated worker without managed entry/runtime files. AOT metadata itself may
# still contain the original managed module's name, so it is not a file dependency.
$bytes = [IO.File]::ReadAllBytes((Join-Path $worker 'QuickExcel.ScriptWorker.exe'))
$pe = [BitConverter]::ToInt32($bytes, 0x3c)
if ([BitConverter]::ToUInt32($bytes, $pe) -ne 0x00004550 -or [BitConverter]::ToUInt16($bytes, $pe + 4) -ne 0x8664) { throw 'Worker is not an x64 PE executable.' }
$optional = $pe + 24
if ([BitConverter]::ToUInt16($bytes, $optional) -ne 0x20b -or [BitConverter]::ToUInt32($bytes, $optional + 112 + 14 * 8) -ne 0) { throw 'Worker is not a native PE32+ executable.' }
$runtime = Get-Content -LiteralPath (Join-Path $package 'QuickExcel.Avalonia.runtimeconfig.json') -Raw | ConvertFrom-Json
if ($runtime.runtimeOptions.configProperties.'System.Text.Json.JsonSerializer.IsReflectionEnabledByDefault' -ne $false) { throw 'Desktop JSON reflection must be disabled in the production package.' }
Write-Host 'PASS: Native AOT worker and one copy of JSC/ICU/TypeScript.'
