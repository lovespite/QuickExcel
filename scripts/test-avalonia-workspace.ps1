[CmdletBinding()]
param([switch]$SkipBuild, [ValidateSet('PasteOptions', 'Workspace', 'Sheets', 'DataTools', 'Borders', 'TextRotation', 'NumberFormat', 'CellStyles', 'Menus', 'ConditionalFormat', 'Colors', 'Grid', 'Drawing', 'PageBreak', 'Compatibility', 'RibbonFilter', 'Automation', 'Execution')][string]$Mode = 'Workspace')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/workspace/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $run | Out-Null
Push-Location $root
try {
    if (-not $SkipBuild) {
        dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -p:RestoreLockedMode=true -v minimal 2>&1 | Tee-Object -FilePath (Join-Path $run 'build.log')
        if ($LASTEXITCODE -ne 0) { throw 'Avalonia workspace build failed.' }
    }
    $exe = Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/QuickExcel.Avalonia.exe'
    $argument = switch ($Mode) { 'PasteOptions' { '--paste-options-smoke' }; 'Sheets' { '--sheets-smoke' }; 'DataTools' { '--data-tools-smoke' }; 'Borders' { '--borders-smoke' }; 'TextRotation' { '--text-rotation-smoke' }; 'NumberFormat' { '--number-format-smoke' }; 'Menus' { '--menus-smoke' }; 'ConditionalFormat' { '--conditional-format-smoke' }; 'Colors' { '--colors-smoke' }; 'Drawing' { '--drawing-smoke' }; 'Grid' { '--grid-smoke' }; 'PageBreak' { '--page-break-smoke' }; 'Compatibility' { '--compatibility-smoke' }; 'RibbonFilter' { '--ribbon-filter-smoke' }; 'Automation' { '--automation-smoke' }; 'Execution' { '--execution-smoke' }; default { '--workspace-smoke' } }
    if ($Mode -eq 'CellStyles') { $argument = '--cell-styles-smoke' }
    $process = Start-Process -FilePath $exe -ArgumentList @($argument, ('"{0}"' -f $run)) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.log') -RedirectStandardError (Join-Path $run 'stderr.log')
    try {
        if (-not $process.WaitForExit(60000)) { throw "Avalonia workspace exceeded 60 seconds: $run" }
        $process.WaitForExit()
        [ordered]@{ exitCode = $process.ExitCode; sourceHead = (& git rev-parse HEAD); dirty = [bool](& git status --porcelain); binarySha256 = (Get-FileHash $exe -Algorithm SHA256).Hash } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'process.json') -Encoding utf8
        $resultPath = Join-Path $run 'result.json'
        if (-not (Test-Path -LiteralPath $resultPath)) { throw "Missing workspace result; exit $($process.ExitCode): $run" }
        $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "Workspace failed: $($result.error) Evidence: $run" }
        Write-Host "PASS: $($result.checks.Count) Avalonia workspace checks. Evidence: $run"
    }
    finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
}
finally { Pop-Location }
