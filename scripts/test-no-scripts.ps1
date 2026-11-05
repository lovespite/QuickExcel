[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/no-scripts/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $run | Out-Null
$previousEnvironment = $env:QUICKEXCEL_NO_SCRIPTS
Push-Location $root
try {
    if (-not $SkipBuild) {
        dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -v minimal
        if ($LASTEXITCODE -ne 0) { throw 'No-scripts smoke build failed.' }
    }
    $exe = Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/QuickExcel.Avalonia.exe'
    foreach ($case in @('setting', 'command-first', 'command-last', 'environment', 'combined', 'saved-setting')) {
        $directory = Join-Path $run $case
        New-Item -ItemType Directory -Path $directory | Out-Null
        if ($case -eq 'saved-setting') {
            New-Item -ItemType Directory -Path (Join-Path $directory 'settings') | Out-Null
            '{"DisableScripts":true}' | Set-Content -LiteralPath (Join-Path $directory 'settings/settings.json') -Encoding utf8
        }
        $env:QUICKEXCEL_NO_SCRIPTS = if ($case -in @('environment', 'combined')) { '1' } else { '0' }
        $arguments = @('--no-scripts-smoke', ('"{0}"' -f $directory))
        if ($case -in @('command-first', 'combined')) { $arguments = @('--no-scripts') + $arguments }
        if ($case -eq 'command-last') { $arguments += '--no-scripts' }
        $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $directory 'stdout.log') -RedirectStandardError (Join-Path $directory 'stderr.log')
        try {
            if (-not $process.WaitForExit(60000)) { throw "No-scripts smoke timed out: $directory" }
            $process.WaitForExit()
            $result = Get-Content -LiteralPath (Join-Path $directory 'result.json') -Raw | ConvertFrom-Json
            if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "No-scripts smoke failed ($case): $($result.error) Evidence: $directory" }
            Write-Host "PASS: $case ($($result.checks.Count) checks)."
        }
        finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
    }
    Write-Host "Evidence: $run"
}
finally { $env:QUICKEXCEL_NO_SCRIPTS = $previousEnvironment; Pop-Location }
