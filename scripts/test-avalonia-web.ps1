[CmdletBinding()]
param([switch]$SkipBuild, [switch]$ScriptsOnly)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/web-tests/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $run | Out-Null
Push-Location $root
try {
    if (-not $SkipBuild) {
        node QuickExcel.Web/build.mjs --tests 2>&1 | Tee-Object -FilePath (Join-Path $run 'web-build.log')
        if ($LASTEXITCODE -ne 0) { throw 'Web test bundle failed.' }
        dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -p:RestoreLockedMode=true -v minimal 2>&1 | Tee-Object -FilePath (Join-Path $run 'build.log')
        if ($LASTEXITCODE -ne 0) { throw 'Avalonia Web host build failed.' }
    }
    $exe = Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/QuickExcel.Avalonia.exe'
    $arguments = @('--web-smoke', ('"{0}"' -f $run)); if ($ScriptsOnly) { $arguments += '--scripts-only' }
    $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.log') -RedirectStandardError (Join-Path $run 'stderr.log')
    try {
        if (-not $process.WaitForExit(180000)) { throw "Avalonia Web smoke exceeded 180 seconds: $run" }
        $process.WaitForExit()
        [ordered]@{ exitCode = $process.ExitCode; sourceHead = (& git rev-parse HEAD); dirty = [bool](& git status --porcelain); binarySha256 = (Get-FileHash $exe -Algorithm SHA256).Hash } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'process.json') -Encoding utf8
        $resultPath = Join-Path $run 'result.json'
        if (-not (Test-Path -LiteralPath $resultPath)) { throw "Missing Web result; exit $($process.ExitCode): $run" }
        $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "Web smoke failed: $($result.error) Evidence: $run" }
        Write-Host "PASS: $($result.checks.Count) Avalonia Web checks. Evidence: $run"
    }
    finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
}
finally { Pop-Location }
