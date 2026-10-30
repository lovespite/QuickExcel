param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$artifactsRoot = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($artifactsRoot, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) {
    throw 'Use a new output directory inside repository artifacts.'
}
[void](New-Item -ItemType Directory -Path $target)
$cases = @(
    @(1234.5678, '#,##0.000'), @(-42.5, '¥#,##0.000;[Red]-¥#,##0.000;"—";@'),
    @(0, '¥#,##0.000;[Red]-¥#,##0.000;"—";@'), @('00123', '0.00;[Red](0.00);"-";"ID: "@'),
    @(2.675, '0.00'), @(-2.675, '0.00'), @(0.125, '0.0%'), @(12200000, '0.0,,"M"'),
    @(12345, '0.00E+00'), @(12345, '##0.0E+0'), @(0.00001234, '0.00E-00'),
    @(999.99, '##0.0E+0'), @(12.3, '???.??'), @(0, '#.##'), @(1, '0.##'),
    @(-12.3, '[<0](0.0);0.00'), @(-12.3, '[>=0]0.00;[Red](0.0)'),
    @(12, '[>100]0.0;[<0]0.00'), @(1234.5, '[$€-407]#,##0.00'),
    @(0, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'),
    @(1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'),
    @(-1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'),
    @(1, '[Green]0.00'), @(-0.0001, '0.00'), @(-0.0001, '0.00;(0.00)'),
    @(1.2, '0.0000000000'), @(0.012, '0.00%'), @(61, '[Blue]yyyy-mm-dd'),
    @(1e20, '0.00E+00'), @(0.5, '.00'), @(0, '???'), @(12, '00000'),
    @(-12.3, '[<0]0.0;0.00'), @(-12.3, '[<0]0.0'), @(-12.3, '[>0]0.0;[<0]0.00;0'),
    @(-12.3, '[>0]0.0;0.00'), @(-12.3, '[>0]0.0;[>10]0.00;0.000'),
    @(12.3, '0.0" kg"'), @(12.3, '"total; "0.00'), @(12.3, '0.0_);[Red](0.0)'),
    @(12.3, '0.00000000000000000000'), @(0, '0.0E+00'), @(0.00000001234, '##0.00E+0')
)
$excel = $null; $book = $null; $sheet = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Number formats'
    $sheet.Columns.Item(1).ColumnWidth = 60
    $rows = @()
    for ($i = 0; $i -lt $cases.Count; $i++) {
        $cell = $sheet.Cells.Item($i + 1, 1)
        try {
            # PowerShell IDispatch uses the host LCID even for NumberFormat (not NumberFormatLocal).
            $format = $cases[$i][1]
            if ($excel.LanguageSettings.LanguageID(2) -eq 2052) {
                foreach ($pair in @(@('Red','红色'), @('Blue','蓝色'), @('Green','绿色'))) {
                    $format = $format.Replace("[$($pair[0])]", "[$($pair[1])]")
                }
            }
            if ($cases[$i][0] -is [string]) { $cell.NumberFormat = '@'; $cell.Value2 = [string]$cases[$i][0] }
            else { $cell.Value2 = [double]$cases[$i][0] }
            $cell.NumberFormat = $format
            $rows += [ordered]@{ Row = $i + 1; Value = $cases[$i][0]; Format = $cases[$i][1]; Text = $cell.Text; Color = $cell.DisplayFormat.Font.Color }
        } catch { throw "Format case $($i + 1) '$($cases[$i][1])': $_" }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
    }
    $book.SaveAs((Join-Path $target 'excel-source.xlsx'), 51)
    $date1904 = $book.Date1904
    $calculation = $excel.Calculation
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    [ordered]@{ Version = $excel.Version; Build = $excel.Build; Language = $excel.LanguageSettings.LanguageID(2); Calculation = $calculation;
        Date1904 = $date1904; SourceSha256 = (Get-FileHash (Join-Path $target 'excel-source.xlsx')).Hash; Cases = $rows } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $target 'excel-results.json') -Encoding utf8
    Write-Output $target
} finally {
    if ($book) { $book.Close($false) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    if ($book) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
