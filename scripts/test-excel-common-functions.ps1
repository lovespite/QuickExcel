param([Parameter(Mandatory)][string]$SpecPath, [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$prefix = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($value) { if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
$spec = Get-Content -LiteralPath $SpecPath -Raw | ConvertFrom-Json
$excel = $null; $book = $null; $sheet = $null; $records = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add(-4167); $sheet = $book.Worksheets.Item(1); $sheet.Name = 'Data'
    foreach ($case in $spec.cases) {
        [void]$sheet.Cells.Clear()
        $cell = $sheet.Range('D6')
        try { $cell.Formula2 = [string]$case.formula } finally { Release-Com $cell }
        $excel.CalculateFullRebuild(); $values = @()
        for ($r = 0; $r -lt $case.rows; $r++) { for ($c = 0; $c -lt $case.columns; $c++) {
            $cell = $sheet.Cells.Item(6 + $r, 4 + $c)
            try {
                $v = $cell.Value2
                $kind = if ($null -eq $v) { 'Blank' } elseif ($v -is [string]) { 'Text' } elseif ($v -is [bool]) { 'Boolean' } elseif ($v -is [double]) { 'Number' } else { 'Error' }
                if ($kind -eq 'Error') { $v = [string]$cell.Text }
                $values += @{ kind = $kind; value = $v }
            } finally { Release-Com $cell }
        } }
        $records += @{ formula = $case.formula; rows = $case.rows; columns = $case.columns; values = $values }
    }
    @{ excelVersion = $excel.Version; excelBuild = $excel.Build; specSha256 = (Get-FileHash -LiteralPath $SpecPath -Algorithm SHA256).Hash; cases = $records } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $target 'excel-common-functions.json') -Encoding utf8
    Write-Host "PASS: captured $($records.Count) independent Excel formula cases in $target"
}
finally { if ($null -ne $book) { $book.Close($false) }; if ($null -ne $excel) { $excel.Quit() }; Release-Com $sheet; Release-Com $book; Release-Com $excel }
