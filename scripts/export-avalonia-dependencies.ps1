[CmdletBinding()]
param([string]$OutputDirectory = 'artifacts/avalonia/dependencies')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts = Join-Path $root 'artifacts'
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
if (-not $output.StartsWith($artifacts + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputDirectory must be inside repository artifacts.'
}
New-Item -ItemType Directory -Force $output | Out-Null
$lock = Get-Content (Join-Path $root 'tools/QuickExcel.AvaloniaProbe/packages.lock.json') -Raw | ConvertFrom-Json -AsHashtable
$assets = Get-Content (Join-Path $root 'artifacts/avalonia/build/obj/QuickExcel.AvaloniaProbe/project.assets.json') -Raw | ConvertFrom-Json -AsHashtable
$archives = [Collections.Generic.List[string]]::new()
$packages = foreach ($entry in $lock.dependencies['net10.0'].GetEnumerator() | Sort-Object Key) {
    if ($entry.Value.type -eq 'Project') { continue }
    $id = $entry.Key
    if ($id -match '(?i)(Xpf|Trial|Commercial|AvaloniaUI\.Pro)') { throw "Disallowed dependency: $id" }
    $version = $entry.Value.resolved
    $relative = "$($id.ToLowerInvariant())/$version"
    $directory = $assets.packageFolders.Keys | ForEach-Object { Join-Path $_ $relative } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $directory) { throw "Restore the probe first: $id/$version" }
    $package = Join-Path $directory "$($id.ToLowerInvariant()).$version.nupkg"
    $bytes = [IO.File]::ReadAllBytes($package)
    $sha512 = [Convert]::ToBase64String([Security.Cryptography.SHA512]::HashData($bytes))
    # NuGet's normalized content hash excludes signature differences; archive SHA-512 is distinct.
    $cacheMetadata = Get-Content (Join-Path $directory '.nupkg.metadata') -Raw | ConvertFrom-Json
    $archiveChecksum = (Get-Content ($package + '.sha512') -Raw).Trim()
    if ($sha512 -cne $archiveChecksum -or $cacheMetadata.contentHash -cne $entry.Value.contentHash) { throw "NuGet identity mismatch: $id/$version" }
    $archives.Add($package)
    [xml]$spec = Get-Content (Join-Path $directory "$($id.ToLowerInvariant()).nuspec")
    $metadata = $spec.SelectSingleNode('/*[local-name()="package"]/*[local-name()="metadata"]')
    $license = $metadata.SelectSingleNode('*[local-name()="license"]')
    $repository = $metadata.SelectSingleNode('*[local-name()="repository"]')
    $licenseUrl = $metadata.SelectSingleNode('*[local-name()="licenseUrl"]')
    [ordered]@{
        id = $id; version = $version; kind = $entry.Value.type
        contentHash = $entry.Value.contentHash; archiveSha512 = $sha512; sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
        source = if ($repository) { $repository.GetAttribute('url') } else { $null }
        commit = if ($repository) { $repository.GetAttribute('commit') } else { $null }
        license = if ($license) { $license.InnerText } else { $null }
        licenseType = if ($license) { $license.GetAttribute('type') } else { $null }
        licenseUrl = if ($licenseUrl) { $licenseUrl.InnerText } else { $null }
        archive = "https://api.nuget.org/v3-flatcontainer/$relative/$($id.ToLowerInvariant()).$version.nupkg"
    }
}
& dotnet nuget verify @archives --all 2>&1 | Set-Content (Join-Path $output 'nuget-signatures.log') -Encoding utf8
if ($LASTEXITCODE -ne 0) { throw "NuGet signature verification failed: $output" }
$npmLock = Get-Content (Join-Path $root 'QuickExcel.Web/package-lock.json') -Raw | ConvertFrom-Json -AsHashtable
$npmPackages = foreach ($entry in $npmLock.packages.GetEnumerator() | Sort-Object Key) {
    if ($entry.Key -eq '') { continue }
    $p = $entry.Value
    [ordered]@{ name = $entry.Key; version = $p['version']; integrity = $p['integrity']; url = $p['resolved']; license = $p['license']; development = [bool]$p['dev']; optional = [bool]$p['optional'] }
}
[ordered]@{ schemaVersion = 1; sourceHead = (& git -C $root rev-parse HEAD).Trim(); nuget = @($packages); npm = @($npmPackages) } |
    ConvertTo-Json -Depth 10 | Set-Content (Join-Path $output 'dependencies.json') -Encoding utf8
Write-Host "Verified $(@($packages).Count) NuGet content hashes; recorded $(@($npmPackages).Count) npm lock identities: $output"
