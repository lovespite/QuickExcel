[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts=Join-Path $root 'artifacts'
$package=[IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($PackageDirectory)) {$PackageDirectory} else {Join-Path $root $PackageDirectory}))
if (-not $package.StartsWith($artifacts+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Package must stay under repository artifacts.' }
$check=$package
while ($check.Length -ge $artifacts.Length) {
    if ((Get-Item -LiteralPath $check -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Package cannot traverse a reparse point.' }
    $check=Split-Path $check -Parent
}
$manifest=Join-Path $package 'release.json'
$identity=Get-Content -LiteralPath $manifest -Raw|ConvertFrom-Json
if ($identity.schemaVersion -ne 1 -or $identity.product -notin @('QuickExcel Avalonia','QuickExcel WPF rollback')) { throw 'Unknown package format.' }
if ($identity.product -eq 'QuickExcel Avalonia' -and ($identity.checks -is [string] -or -not $identity.checks.passed)) { throw 'Selecting Avalonia requires a passed production self-check.' }
foreach($file in $identity.files) {
    $path=[IO.Path]::GetFullPath((Join-Path $package $file.path))
    if (-not $path.StartsWith($package+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $file.sha256) { throw "Package file identity mismatch: $($file.path)" }
}
$pointer=Join-Path $artifacts 'avalonia/active-windows.json'; $temporary=$pointer+'.'+[Guid]::NewGuid().ToString('N')+'.tmp'
[ordered]@{schemaVersion=1;directory=$package;manifestSha256=(Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content -LiteralPath $temporary -Encoding utf8
Move-Item -LiteralPath $temporary -Destination $pointer -Force
Write-Host "Selected $($identity.product): $package"
