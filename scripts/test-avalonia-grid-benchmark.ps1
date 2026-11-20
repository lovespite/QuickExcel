[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/benchmarks/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $run | Out-Null
Push-Location $root
try {
    if (-not $SkipBuild) {
        dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -p:RestoreLockedMode=true -v minimal 2>&1 | Tee-Object -FilePath (Join-Path $run 'build.log')
        if ($LASTEXITCODE -ne 0) { throw 'Avalonia benchmark build failed.' }
    }
    $exe = Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/quickxl.exe'
    $identity = [ordered]@{ sourceHead = (& git rev-parse HEAD); dirty = [bool](& git status --porcelain); binarySha256 = (Get-FileHash $exe -Algorithm SHA256).Hash }
    $identity | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'identity.json') -Encoding utf8
    $process = Start-Process -FilePath $exe -ArgumentList @('--grid-benchmark', ('"{0}"' -f $run)) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.log') -RedirectStandardError (Join-Path $run 'stderr.log')
    try {
        if (-not $process.WaitForExit(60000)) { throw "Avalonia benchmark exceeded 60 seconds: $run" }
        $process.WaitForExit()
        [ordered]@{ exitCode = $process.ExitCode } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'process.json') -Encoding utf8
        $resultPath = Join-Path $run 'result.json'
        if (-not (Test-Path -LiteralPath $resultPath)) { throw "Missing benchmark result; exit $($process.ExitCode): $run" }
        $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "Benchmark failed: $($result.error) Evidence: $run" }
        Write-Host "PASS: Avalonia scroll, drag and load budgets. Evidence: $run"
    }
    finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
}
finally { Pop-Location }
