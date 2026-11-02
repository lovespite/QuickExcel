[CmdletBinding()]
param([switch]$SkipBuild, [ValidateSet('all', 'editor', 'pdf', 'pdf-visual', 'cold-editor', 'cold-pdf', 'retention', 'concurrent')][string]$Mode = 'all')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/web-benchmarks/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $run | Out-Null
Push-Location $root
try {
    if (-not $SkipBuild) {
        node QuickExcel.Web/build.mjs --tests 2>&1 | Tee-Object -FilePath (Join-Path $run 'web-build.log')
        if ($LASTEXITCODE -ne 0) { throw 'Web benchmark bundle failed.' }
        dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -p:RestoreLockedMode=true -v minimal 2>&1 | Tee-Object -FilePath (Join-Path $run 'build.log')
        if ($LASTEXITCODE -ne 0) { throw 'Avalonia benchmark build failed.' }
    }
    if ($Mode -eq 'all') {
        # Separate process trees and profiles for distinct load budgets; never carry an
        # editor benchmark's retained managed heap into the ordinary PDF baseline.
        $modes = @('editor', 'pdf', 'retention', 'concurrent') + (@('cold-editor') * 5) + (@('cold-pdf') * 5)
        foreach ($scenario in $modes) { & $PSCommandPath -SkipBuild -Mode $scenario *>&1 | Tee-Object -FilePath (Join-Path $run 'runs.log') -Append }
        [ordered]@{ passed = $true; scenarios = $modes; note = 'Each scenario ran in its own process/profile; see runs.log for evidence directories.' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'result.json') -Encoding utf8
        Write-Host "PASS: all independent Web benchmarks. Evidence: $run"
        return
    }
    $exe = Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/QuickExcel.Avalonia.exe'
    $process = Start-Process -FilePath $exe -ArgumentList @('--web-benchmark', ('"{0}"' -f $run), $Mode) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.log') -RedirectStandardError (Join-Path $run 'stderr.log')
    try {
        if (-not $process.WaitForExit(600000)) { throw "Web benchmark exceeded ten minutes: $run" }
        $process.WaitForExit()
        [ordered]@{ mode = $Mode; exitCode = $process.ExitCode; sourceHead = (& git rev-parse HEAD); dirty = [bool](& git status --porcelain); binarySha256 = (Get-FileHash $exe -Algorithm SHA256).Hash } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'process.json') -Encoding utf8
        [ordered]@{
            cpu = @(Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors)
            graphics = @(Get-CimInstance Win32_VideoController | Select-Object Name, DriverVersion, CurrentHorizontalResolution, CurrentVerticalResolution)
            fonts = @(Get-ChildItem -LiteralPath (Join-Path $env:WINDIR 'Fonts') -File | Get-FileHash -Algorithm SHA256 | Select-Object Path, Hash)
            note = 'System font inventory and graphics driver identity; CPU submission and browser animation-frame observations do not measure physical presentation.'
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $run 'machine.json') -Encoding utf8
        $result = Get-Content -LiteralPath (Join-Path $run 'result.json') -Raw | ConvertFrom-Json
        if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "Web benchmark failed. Evidence: $run" }
        Write-Host "PASS: Web benchmark. Evidence: $run"
    }
    finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
}
finally { Pop-Location }
