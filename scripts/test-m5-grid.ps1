[CmdletBinding()]
param(
    [string]$PackageDirectory = 'artifacts\QuickExcel',
    [ValidateSet('Interaction', 'Workspace', 'Product', 'Ribbon', 'DataOperations', 'GridIntegration', 'FastGridBenchmark', 'FastGridDragBenchmark', 'Benchmark', 'ResizeBenchmark', 'All')][string]$Mode = 'All'
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts = [IO.Path]::GetFullPath((Join-Path $root 'artifacts'))
$package = [IO.Path]::GetFullPath((Join-Path $root $PackageDirectory))
if (-not $package.StartsWith($artifacts + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'PackageDirectory must be under this repository artifacts directory.'
}
$exe = Join-Path $package 'QuickExcel.exe'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Missing executable: $exe" }
$run = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$evidence = Join-Path $artifacts "validation\m5-grid\$run"
New-Item -ItemType Directory -Path $evidence | Out-Null
$modes = if ($Mode -eq 'All') { @('Interaction', 'Workspace', 'Product', 'Ribbon', 'Benchmark') } else { @($Mode) }
foreach ($scenario in $modes) {
    $directory = Join-Path $evidence $scenario
    New-Item -ItemType Directory -Path $directory | Out-Null
    $argument = switch ($scenario) { 'GridIntegration' { '--grid-integration-smoke' } 'FastGridDragBenchmark' { '--fast-grid-drag-benchmark' } 'FastGridBenchmark' { '--fast-grid-benchmark' } 'DataOperations' { '--data-operations-smoke' } 'Ribbon' { '--ribbon-smoke' } 'ResizeBenchmark' { '--resize-benchmark' } 'Benchmark' { '--grid-benchmark' } 'Workspace' { '--workspace-smoke' } 'Product' { '--product-smoke' } default { '--interaction-smoke' } }
    $result = switch ($scenario) { 'GridIntegration' { 'grid-integration-smoke-result.txt' } 'FastGridDragBenchmark' { 'fast-grid-drag-benchmark-result.txt' } 'FastGridBenchmark' { 'fast-grid-benchmark-result.txt' } 'DataOperations' { 'data-operations-smoke-result.txt' } 'Ribbon' { 'ribbon-smoke-result.txt' } 'ResizeBenchmark' { 'resize-benchmark-result.txt' } 'Benchmark' { 'grid-benchmark-result.txt' } 'Workspace' { 'workspace-smoke-result.txt' } 'Product' { 'product-smoke-result.txt' } default { 'interaction-smoke-result.txt' } }
    $process = $null
    try {
        $process = Start-Process -FilePath $exe -ArgumentList @($argument, ('"{0}"' -f $directory)) -WorkingDirectory $package -WindowStyle Hidden -PassThru
        if (-not $process.WaitForExit(120000)) { throw "M5 $scenario exceeded 120 seconds. Evidence: $directory" }
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "M5 $scenario failed with exit code $($process.ExitCode). Evidence: $directory" }
        $resultPath = Join-Path $directory $result
        if (-not (Test-Path -LiteralPath $resultPath) -or -not (Get-Content -LiteralPath $resultPath -Raw).StartsWith('PASS:')) {
            throw "Missing fresh M5 result: $directory"
        }
        Get-Content -LiteralPath $resultPath
    }
    finally {
        if ($null -ne $process) {
            if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
            $process.Dispose()
        }
    }
}
$processor = $null
$physicalMemory = $null
$environmentNotes = @()
try {
    $processor = Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name
    $physicalMemory = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory
}
catch { $environmentNotes += 'CIM hardware inventory unavailable: ' + $_.Exception.Message }
[ordered]@{
    sourceHead = (& git -C $root rev-parse HEAD | Out-String).Trim()
    sourceDirty = [bool]((& git -C $root status --porcelain | Out-String).Trim())
    package = $package
    binarySha256 = (Get-FileHash -LiteralPath (Join-Path $package 'QuickExcel.dll') -Algorithm SHA256).Hash
    os = [Environment]::OSVersion.ToString()
    processor = $processor
    physicalMemoryBytes = $physicalMemory
    environmentNotes = $environmentNotes
    evidence = $evidence
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'run.json') -Encoding utf8
Write-Host "M5 evidence: $evidence"
