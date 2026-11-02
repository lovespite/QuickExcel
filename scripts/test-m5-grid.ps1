[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$PackageDirectory,
    [ValidateSet('Interaction', 'Workspace', 'Product', 'Ribbon', 'Reporting', 'DataOperations', 'FormatEfficiency', 'GridIntegration', 'FastGridBenchmark', 'HorizontalScrollBenchmark', 'FastGridDragBenchmark', 'Benchmark', 'ResizeBenchmark', 'All')][string]$Mode = 'All',
    [string]$WorkbookPath,
    [switch]$SoftwareRendering,
    [switch]$DefaultRendering
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($SoftwareRendering -and $DefaultRendering) { throw 'Choose only one rendering override.' }
if ($WorkbookPath) {
    if ($Mode -notin @('HorizontalScrollBenchmark', 'Reporting')) { throw 'WorkbookPath is supported only by HorizontalScrollBenchmark and Reporting.' }
    $WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw 'WorkbookPath must be a file.' }
}
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
$evidence = Join-Path $artifacts "validation\m5-grid\$run"
New-Item -ItemType Directory -Path $evidence | Out-Null
$modes = if ($Mode -eq 'All') { @('Interaction', 'Workspace', 'Product', 'Ribbon', 'Benchmark') } else { @($Mode) }
$validationFailures = [Collections.Generic.List[string]]::new()
foreach ($scenario in $modes) {
    $directory = Join-Path $evidence $scenario
    New-Item -ItemType Directory -Path $directory | Out-Null
    $argument = switch ($scenario) { 'FormatEfficiency' { '--format-efficiency-smoke' } 'HorizontalScrollBenchmark' { '--horizontal-scroll-benchmark' } 'GridIntegration' { '--grid-integration-smoke' } 'FastGridDragBenchmark' { '--fast-grid-drag-benchmark' } 'FastGridBenchmark' { '--fast-grid-benchmark' } 'DataOperations' { '--data-operations-smoke' } 'Ribbon' { '--ribbon-smoke' } 'ResizeBenchmark' { '--resize-benchmark' } 'Benchmark' { '--grid-benchmark' } 'Workspace' { '--workspace-smoke' } 'Product' { '--product-smoke' } default { '--interaction-smoke' } }
    $result = switch ($scenario) { 'FormatEfficiency' { 'format-efficiency-result.txt' } 'HorizontalScrollBenchmark' { 'horizontal-scroll-benchmark-result.txt' } 'GridIntegration' { 'grid-integration-smoke-result.txt' } 'FastGridDragBenchmark' { 'fast-grid-drag-benchmark-result.txt' } 'FastGridBenchmark' { 'fast-grid-benchmark-result.txt' } 'DataOperations' { 'data-operations-smoke-result.txt' } 'Ribbon' { 'ribbon-smoke-result.txt' } 'ResizeBenchmark' { 'resize-benchmark-result.txt' } 'Benchmark' { 'grid-benchmark-result.txt' } 'Workspace' { 'workspace-smoke-result.txt' } 'Product' { 'product-smoke-result.txt' } default { 'interaction-smoke-result.txt' } }
    $process = $null
    if ($scenario -eq 'Reporting') { $argument = '--reporting-smoke'; $result = 'reporting-result.txt' }
    try {
        $arguments = @($argument, ('"{0}"' -f $directory))
        if ($WorkbookPath) { $arguments += ('"{0}"' -f $WorkbookPath) }
        if ($SoftwareRendering) { $arguments += '--software-rendering' }
        if ($DefaultRendering) { $arguments += '--default-rendering' }
        $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $package -WindowStyle Hidden -PassThru
        if (-not $process.WaitForExit(120000)) { throw "M5 $scenario exceeded 120 seconds. Evidence: $directory" }
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "M5 $scenario failed with exit code $($process.ExitCode). Evidence: $directory" }
        $resultPath = Join-Path $directory $result
        if (-not (Test-Path -LiteralPath $resultPath)) {
            throw "Missing fresh M5 result: $directory"
        }
        $summary = Get-Content -LiteralPath $resultPath -TotalCount 1
        if (-not $summary.StartsWith('PASS:')) { $validationFailures.Add("M5 $scenario did not pass: $summary Evidence: $directory") }
        $rendering = Get-Content -LiteralPath (Join-Path $directory 'rendering.json') -Raw | ConvertFrom-Json
        $expectedMode = if ($SoftwareRendering) { 'SoftwareOnly' } else { 'Default' }
        $expectedSelection = if ($SoftwareRendering) { 'explicit-software' } elseif ($DefaultRendering) { 'explicit-default' } else { 'automatic-default' }
        if ($rendering.processRenderMode -ne $expectedMode -or $rendering.renderModeSelection -ne $expectedSelection) {
            throw "Rendering mode was not applied ($expectedSelection): $directory"
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
    requestedRenderMode = if ($SoftwareRendering) { 'SoftwareOnly' } else { 'Default' }
    renderModeSelection = if ($SoftwareRendering) { 'explicit-software' } elseif ($DefaultRendering) { 'explicit-default' } else { 'automatic-default' }
    binarySha256 = (Get-FileHash -LiteralPath (Join-Path $package 'QuickExcel.dll') -Algorithm SHA256).Hash
    os = [Environment]::OSVersion.ToString()
    processor = $processor
    physicalMemoryBytes = $physicalMemory
    environmentNotes = $environmentNotes
    evidence = $evidence
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'run.json') -Encoding utf8
Write-Host "M5 evidence: $evidence"
if ($validationFailures.Count -gt 0) { throw ($validationFailures -join [Environment]::NewLine) }
