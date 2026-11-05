param([Parameter(Mandatory)][string]$OutputDirectory, [Parameter(Mandatory)][string]$QuickExcelCsv)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$prefix = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($value) { if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
$excel = $null; $book = $null; $sheet = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    $opened = $excel.Workbooks.Open([IO.Path]::GetFullPath($QuickExcelCsv, $repo), 0, $true)
    Release-Com $opened
    $book = $excel.ActiveWorkbook; $sheet = $book.Worksheets.Item(1)
    $cell = $sheet.Cells.Item(1, 1)
    try { if ($cell.Value2 -cne "中文,`"甲`"`n乙") { throw ('QuickExcel CSV -> Excel differs: ' + ($cell.Value2 | ConvertTo-Json -Compress)) } } finally { Release-Com $cell }
    $book.Close($false); Release-Com $sheet; $sheet = $null; Release-Com $book; $book = $null
    $book = $excel.Workbooks.Add(-4167); $sheet = $book.Worksheets.Item(1)
    $values = @(@('001234567890123456789', '=1+1', "中文,`"引号`"`n第二行"), @('0002', '', '结束'))
    for ($r = 0; $r -lt $values.Count; $r++) {
        for ($c = 0; $c -lt 3; $c++) {
            $cell = $sheet.Cells.Item($r + 1, $c + 1)
            try { $cell.NumberFormat = '@'; $cell.Value2 = [string]$values[$r][$c] } finally { Release-Com $cell }
        }
    }
    $csv = Join-Path $target 'excel-output.csv'; $book.SaveAs($csv, 62)
    $book.Close($false); Release-Com $sheet; $sheet = $null; Release-Com $book; $book = $null
    [ordered]@{ excelVersion = $excel.Version; excelBuild = $excel.Build; sourceSha256 = (Get-FileHash -LiteralPath $QuickExcelCsv -Algorithm SHA256).Hash;
        csvSha256 = (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash; text = [IO.File]::ReadAllText($csv); rows = $values; quickExcelToExcelPassed = $true } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $target 'text-exchange-excel.json') -Encoding utf8
    Write-Host "PASS: QuickExcel CSV opened in Excel; independent Excel CSV and expected text captured in $target"
}
finally { if ($null -ne $book) { $book.Close($false) }; if ($null -ne $excel) { $excel.Quit() }; Release-Com $sheet; Release-Com $book; Release-Com $excel }
