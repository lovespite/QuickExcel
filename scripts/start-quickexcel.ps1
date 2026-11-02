param()
$ApplicationArguments=@($args)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$pointer=Join-Path $root 'artifacts/avalonia/active-windows.json'
if (Test-Path -LiteralPath $pointer) {
    $active=Get-Content -LiteralPath $pointer -Raw | ConvertFrom-Json
    $package=[IO.Path]::GetFullPath($active.directory)
    $prefix=(Join-Path $root 'artifacts')+[IO.Path]::DirectorySeparatorChar
    if (-not $package.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) { throw 'Selected package must stay under repository artifacts.' }
    $manifest=Join-Path $package 'release.json'
    if ((Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash -ne $active.manifestSha256) { throw 'Selected package identity changed; select a verified package again.' }
    $identity=Get-Content -LiteralPath $manifest -Raw|ConvertFrom-Json
    $executable=if ($identity.product -eq 'QuickExcel WPF rollback') { 'QuickExcel.exe' } else { 'QuickExcel.Avalonia.exe' }
    $entry=@($identity.files|Where-Object {$_.path -eq $executable})
    if ($entry.Count -ne 1 -or (Get-FileHash -LiteralPath (Join-Path $package $executable) -Algorithm SHA256).Hash -ne $entry[0].sha256) { throw 'Selected executable identity mismatch.' }
    function Quote-Argument([string]$value) {
        $escaped=[regex]::Replace($value,'(\\*)"','$1$1\"')
        $escaped=[regex]::Replace($escaped,'(\\+)$','$1$1')
        return '"'+$escaped+'"'
    }
    $launch=@{FilePath=(Join-Path $package $executable);WorkingDirectory=$package;PassThru=$true;Wait=$true}
    if($ApplicationArguments.Count){$launch.ArgumentList=($ApplicationArguments|ForEach-Object {Quote-Argument $_}) -join ' '}
    if($ApplicationArguments.Count -and $ApplicationArguments[0] -in @('--self-check','--workspace-smoke')){$launch.WindowStyle='Hidden'}
    $process=Start-Process @launch
    try {exit $process.ExitCode} finally {$process.Dispose()}
}
Push-Location $root
try {
    & "$PSScriptRoot/build-avalonia-web.ps1"
    dotnet run --project QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -- @ApplicationArguments
    exit $LASTEXITCODE
} finally { Pop-Location }
