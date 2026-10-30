param([Parameter(Mandatory)][string]$WorkbookPath, [Parameter(Mandatory)][string]$ReferencePath,
    [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$artifactsRoot = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($artifactsRoot, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) {
    throw 'Use a new output directory inside repository artifacts.'
}
$inputPath = [IO.Path]::GetFullPath($WorkbookPath, $repo)
$before = (Get-FileHash -LiteralPath $inputPath).Hash
$reference = Get-Content -LiteralPath $ReferencePath -Raw | ConvertFrom-Json
[void](New-Item -ItemType Directory -Path $target)
$excel = $null; $book = $null; $sheet = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($inputPath, 0, $true)
    $sheet = $book.Worksheets.Item(1)
    $results = @()
    foreach ($case in $reference.Cases) {
        $cell = $sheet.Cells.Item($case.Row, 1)
        try {
            $expected = [string]$case.Text; $actual = [string]$cell.Text
            if ($case.Format.Contains('* ')) {
                # Compare all content except the width-dependent currency fill run.
                $expected = [regex]::Replace($expected, '(?<=\$) +', '')
                $actual = [regex]::Replace($actual, '(?<=\$) +', '')
            }
            if ($expected -match '^#+$') { $expected = '#'; if ($actual -match '^#+$') { $actual = '#' } }
            $value = $cell.Value2
            $passed = $expected -ceq $actual -and $case.Color -eq $cell.DisplayFormat.Font.Color -and $case.Value -ceq $value
            $results += [ordered]@{ Row=$case.Row; Passed=$passed; Expected=$expected; Actual=$actual; Value=$value; Color=$cell.DisplayFormat.Font.Color }
        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
    }
    # Workbooks.Open defaults to xlNormalLoad; object-model opens do not attempt recovery.
    $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book); $book=$null
    $unchanged = $before -eq (Get-FileHash -LiteralPath $inputPath).Hash
    [ordered]@{ Version=$excel.Version; Build=$excel.Build; Language=$excel.LanguageSettings.LanguageID(2);
        CorruptLoad='xlNormalLoad'; SourceUnchanged=$unchanged; WorkbookSha256=$before; Results=$results } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $target 'excel-reopen.json') -Encoding utf8
    if (-not $unchanged -or @($results | Where-Object { -not $_.Passed }).Count -gt 0) { throw 'Excel reopen comparison failed; see evidence.' }
    Write-Output "$($results.Count) Excel reopen checks passed; normal load; source unchanged=$unchanged"
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
