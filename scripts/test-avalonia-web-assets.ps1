[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/web-asset-tests/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$build = Join-Path $run 'build'
$publish = Join-Path $run 'publish'
$project = Join-Path $root 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj'
$source = Join-Path $root 'artifacts/avalonia/web'
$checks = [Collections.Generic.List[string]]::new()
New-Item -ItemType Directory -Path $run | Out-Null

function Verify-Bundle([string]$directory) {
    $manifestPath = Join-Path $directory 'manifest.json'
    if ((Get-FileHash -LiteralPath $manifestPath).Hash -ne (Get-FileHash -LiteralPath (Join-Path $source 'manifest.json')).Hash) {
        throw "Copied Web manifest differs from the source: $directory"
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    foreach ($asset in $manifest.files) {
        $file = Join-Path $directory $asset.name
        if ((Get-Item -LiteralPath $file).Length -ne $asset.size -or (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne $asset.sha256) {
            throw "Copied Web asset differs from the manifest: $file"
        }
    }
}

function Change-NewerOutput([string]$directory) {
    # Reproduce changed HTML with a newer timestamp, including the original
    # missing-ellipsis failure. Also cover same-size changes and the manifest.
    foreach ($name in @('script.html', 'preview.html', 'manifest.json')) {
        $file = Join-Path $directory $name
        [IO.File]::AppendAllText($file, "`n ")
        [IO.File]::SetLastWriteTimeUtc($file, [DateTime]::UtcNow.AddHours(1))
    }
    $file = Join-Path $directory 'script.js'
    $bytes = [IO.File]::ReadAllBytes($file); $bytes[0] = $bytes[0] -bxor 1
    [IO.File]::WriteAllBytes($file, $bytes)
    [IO.File]::SetLastWriteTimeUtc($file, [DateTime]::UtcNow.AddHours(1))
    $rejected = $false
    try { Verify-Bundle $directory } catch { $rejected = $true }
    if (-not $rejected) { throw 'The changed-output fixture was not detected.' }
    Verify-Bundle $source
}

function Build-Host([string]$label) {
    dotnet build $project -c Release --artifacts-path $build -p:RestoreLockedMode=true -v minimal 2>&1 | Tee-Object -FilePath (Join-Path $run ($label + '.log'))
    if ($LASTEXITCODE -ne 0) { throw "Host build failed: $label" }
}
function Publish-Host([string]$label) {
    dotnet publish $project -c Release --no-restore --self-contained false --artifacts-path $build -o $publish -v minimal 2>&1 | Tee-Object -FilePath (Join-Path $run ($label + '.log'))
    if ($LASTEXITCODE -ne 0) { throw "Host publish failed: $label" }
}

Push-Location $root
try {
    node QuickExcel.Web/build.mjs
    if ($LASTEXITCODE -ne 0) { throw 'Production Web bundle build failed. Run scripts/build-avalonia-web.ps1 to install the locked dependencies.' }
    Build-Host 'build-before'
    $output = Join-Path $build 'bin/QuickExcel.Desktop.Avalonia/release/Web'
    Verify-Bundle $output
    Change-NewerOutput $output
    Build-Host 'build-after'
    Verify-Bundle $output
    $checks.Add('Incremental build repairs newer changed HTML, same-size JavaScript and manifest copies')

    Publish-Host 'publish-before'
    Verify-Bundle (Join-Path $publish 'Web')
    Change-NewerOutput (Join-Path $publish 'Web')
    Publish-Host 'publish-after'
    Verify-Bundle (Join-Path $publish 'Web')
    $checks.Add('Incremental publish repairs newer changed Web resources and manifest copies')
    [ordered]@{ passed = $true; checks = @($checks) } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'result.json') -Encoding utf8
    Write-Host "PASS: $($checks.Count) Web asset copy checks. Evidence: $run"
}
finally { Pop-Location }
