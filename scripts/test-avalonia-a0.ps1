[CmdletBinding()]
param([switch]$SkipWebBuild, [switch]$Interactive, [switch]$VirtualHttps)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/avalonia/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory $run | Out-Null
Push-Location $root
try {
    if (-not $SkipWebBuild) { & (Join-Path $PSScriptRoot 'build-avalonia-web.ps1') }
    & dotnet build tools/QuickExcel.AvaloniaProbe/QuickExcel.AvaloniaProbe.csproj -c Release --artifacts-path artifacts/avalonia/build -p:RestoreLockedMode=true -v minimal 2>&1 |
        Tee-Object -FilePath (Join-Path $run 'build.log')
    if ($LASTEXITCODE -ne 0) { throw 'A0 build failed.' }
    & (Join-Path $PSScriptRoot 'export-avalonia-dependencies.ps1') -OutputDirectory ([IO.Path]::GetRelativePath($root, (Join-Path $run 'dependencies')))
    $exe = Join-Path $root 'artifacts/avalonia/build/bin/QuickExcel.AvaloniaProbe/release/QuickExcel.AvaloniaProbe.exe'
    $fonts = Join-Path $root 'artifacts/skia-p0/fonts'
    if (-not (Test-Path (Join-Path $fonts 'NotoSansCJKsc-Regular.otf'))) { throw 'Run scripts/get-rendering-test-fonts.ps1 before the A0 probe.' }
    $arguments = @(('"{0}"' -f $run), '--web', ('"{0}"' -f (Join-Path $root 'artifacts/avalonia/web')), '--fonts', ('"{0}"' -f $fonts))
    if ($Interactive) { $arguments += '--interactive' }
    if ($VirtualHttps) { $arguments += '--https' }
    $processOptions = @{ FilePath = $exe; ArgumentList = $arguments; WorkingDirectory = $root; PassThru = $true;
        RedirectStandardOutput = (Join-Path $run 'stdout.log'); RedirectStandardError = (Join-Path $run 'stderr.log') }
    if (-not $Interactive) { $processOptions.WindowStyle = 'Hidden' }
    $process = Start-Process @processOptions
    $completed = $false
    try {
        if ($Interactive) { $process.WaitForExit() }
        elseif (-not $process.WaitForExit(60000)) { throw "A0 probe timed out: $run" }
        $process.WaitForExit()
        $completed = $true
        if ($process.ExitCode -ne 0) { throw "A0 probe failed ($($process.ExitCode)): $run" }
        $result = Get-Content (Join-Path $run 'result.json') -Raw | ConvertFrom-Json
        if (-not $result.passed -or $result.checks.Count -ne 3) { throw "A0 probe incomplete: $run" }
        Write-Host "PASS: automated A0 checks (real IME/DPI/desktop acceptance remains separate): $run"
    }
    finally {
        if (-not $process.HasExited) { $process.Kill($true); [void]$process.WaitForExit(5000) }
        [ordered]@{ schemaVersion = 1; processExited = $completed; exitCode = $process.ExitCode;
            interactive = [bool]$Interactive; virtualHttps = [bool]$VirtualHttps; sourceHead = (& git rev-parse HEAD).Trim();
            sourceChanges = @(& git status --porcelain); utc = [DateTime]::UtcNow.ToString('O') } |
            ConvertTo-Json -Depth 4 | Set-Content (Join-Path $run 'process.json') -Encoding utf8
        $process.Dispose()
    }
}
finally { Pop-Location }
