param()
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
Push-Location $repo
try {
    $evidence = Join-Path $repo ("artifacts/compatibility/e0-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    dotnet test .\tools\QuickExcel.CompatibilityProbe.Tests\QuickExcel.CompatibilityProbe.Tests.csproj -c Release --artifacts-path .\artifacts\build -v minimal --logger trx --results-directory "$evidence-test-results"
    if ($LASTEXITCODE -ne 0) { throw 'XLSX compatibility probe tests failed.' }
    dotnet run --project .\tools\QuickExcel.CompatibilityProbe\QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path .\artifacts\build -- baseline $evidence
    if ($LASTEXITCODE -ne 0) { throw "Baseline generation failed; evidence retained: $evidence" }
    [ordered]@{
        RecordedUtc = [DateTimeOffset]::UtcNow.ToString('o')
        GitHead = (git rev-parse HEAD)
        WorkingTreeDirty = [bool](git status --porcelain)
        DotnetSdk = (dotnet --version)
        OpenXmlSdk = '3.3.0'
        SchemaTarget = 'Office2019'
        ExcelVerification = 'pending-real-excel'
        TestResults = "$evidence-test-results"
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'run-metadata.json') -Encoding utf8
    Write-Host "Evidence: $evidence"
    Write-Host 'Baseline generation is not an Excel compatibility pass. Review summary.md and report.json validation errors, losses and rejected paths.'
}
finally { Pop-Location }
