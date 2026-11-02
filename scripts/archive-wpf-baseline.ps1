[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$IdentityFile)
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent)); $artifacts=Join-Path $root 'artifacts'
$identityPath=[IO.Path]::GetFullPath($IdentityFile)
if (-not $identityPath.StartsWith($artifacts+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Baseline identity must be inside artifacts.' }
$identity=Get-Content -LiteralPath $identityPath -Raw|ConvertFrom-Json
$source=[IO.Path]::GetFullPath($identity.package)
if (-not $source.StartsWith($artifacts+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Baseline package must be inside artifacts.' }
if (@(Get-ChildItem -LiteralPath $source -Recurse -Force | Where-Object {$_.Attributes -band [IO.FileAttributes]::ReparsePoint})) { throw 'Linked baseline content refused.' }
$runId=[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8)
$destination=Join-Path $artifacts "avalonia/rollback/$runId/application"
New-Item -ItemType Directory -Path $destination | Out-Null
foreach($file in $identity.files) {
    $path=[IO.Path]::GetFullPath((Join-Path $source $file.path))
    if (-not $path.StartsWith($source+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $file.sha256) { throw "Frozen WPF baseline changed: $($file.path)" }
    $target=Join-Path $destination $file.path
    New-Item -ItemType Directory -Force (Split-Path $target -Parent)|Out-Null
    Copy-Item -LiteralPath $path -Destination $target
}
if (@(Get-ChildItem -LiteralPath $source -Recurse -File).Count -ne $identity.files.Count) { throw 'Baseline has unrecorded files.' }
$files=@($identity.files|ForEach-Object {[ordered]@{path=$_.path;size=$_.bytes;sha256=$_.sha256}})
[ordered]@{schemaVersion=1;product='QuickExcel WPF rollback';sourceBaseline=$identity.sourceBaseline;originalMetadata=$identity.originalMetadata;baselineIdentitySha256=(Get-FileHash -LiteralPath $identityPath -Algorithm SHA256).Hash;files=$files}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $destination 'release.json') -Encoding utf8
$archive=Join-Path $artifacts "avalonia/rollback/QuickExcel-WPF-baseline-$runId.zip"
Compress-Archive -LiteralPath $destination -DestinationPath $archive
(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash+'  '+[IO.Path]::GetFileName($archive)|Set-Content -LiteralPath ($archive+'.sha256') -Encoding ascii
Write-Host "Archived verified WPF baseline: $destination"
