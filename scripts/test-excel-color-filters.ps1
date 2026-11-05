param([Parameter(Mandatory)][string]$InputDirectory, [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$output = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$inputPath = [IO.Path]::GetFullPath($InputDirectory, $repo)
$prefix = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $output.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $output)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $output)
function Release-Com($value) { if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
function Visible-Rows($sheet) { @(2..5 | Where-Object { $row = $sheet.Rows.Item($_); try { -not $row.Hidden } finally { Release-Com $row } }) }
$excel = $null; $book = $null; $sheet = $null; $cases = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    foreach ($name in @('fill', 'font', 'none')) {
        $expected = @(switch ($name) { fill { @(2,4) }; font { @(3,4) }; none { @(5) } })
        $source = Join-Path $inputPath "$name.xlsx"
        $book = $excel.Workbooks.Open($source, 0, $true); $sheet = $book.Worksheets.Item(1)
        $sheet.AutoFilter.ApplyFilter()
        $actual = @(Visible-Rows $sheet)
        if (($actual -join ',') -ne ($expected -join ',')) { throw "$name reapply differs: $actual" }
        $book.SaveAs((Join-Path $output "roundtrip-$name.xlsx"), 51)
        $book.Close($false); Release-Com $sheet; $sheet = $null; Release-Com $book; $book = $null
        $book = $excel.Workbooks.Add(-4167); $sheet = $book.Worksheets.Item(1)
        $sheet.Range('A1').Value2 = 'Amount'
        for ($r=2; $r -le 5; $r++) { $cell=$sheet.Cells.Item($r,1); try { $cell.Value2=($r-1)*10 } finally { Release-Com $cell } }
        $sheet.Range('A2').Interior.Color=255; $sheet.Range('A3').Interior.Color=16711680; $sheet.Range('A4').Interior.Color=255
        $sheet.Range('A3:A4').Font.Color=255
        $operator = switch ($name) { fill {8}; font {9}; none {12} }
        [void]$sheet.Range('A1:A5').AutoFilter(1,255,$operator)
        $independent=@(Visible-Rows $sheet)
        if (($independent -join ',') -ne ($expected -join ',')) { throw "$name independent result differs: $independent" }
        $produced=Join-Path $output "excel-$name.xlsx"; $book.SaveAs($produced,51)
        $book.Close($false); Release-Com $sheet; $sheet=$null; Release-Com $book; $book=$null
        $zip=[IO.Compression.ZipFile]::OpenRead($produced)
        try {
            $parts=[ordered]@{}
            foreach($entry in $zip.Entries) { if($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')) { $reader=[IO.StreamReader]::new($entry.Open()); try { $parts[$entry.FullName]=$reader.ReadToEnd() } finally {$reader.Dispose()} } }
        } finally {$zip.Dispose()}
        $cases += [ordered]@{ name=$name; visibleRows=$expected; quickExcelSha256=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;
            excelSha256=(Get-FileHash -LiteralPath $produced -Algorithm SHA256).Hash; parts=$parts }
    }
    [ordered]@{ excelVersion=$excel.Version; excelBuild=$excel.Build; cases=$cases } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $output 'excel-color-filters.json') -Encoding utf8
    Write-Host "PASS: color filters reapply in Excel; independent fixtures: $output"
}
finally { if($null -ne $book){$book.Close($false)}; if($null -ne $excel){$excel.Quit()}; Release-Com $sheet; Release-Com $book; Release-Com $excel }
