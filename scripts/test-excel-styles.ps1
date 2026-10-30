param([Parameter(Mandatory)][string]$OutputDirectory, [string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$artifactsRoot=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($artifactsRoot,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $target)
$excel=$null; $book=$null; $sheet=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false; $excel.AutomationSecurity=3
    if ($WorkbookPath) { $book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath,$repo),0,$true) }
    else { $book=$excel.Workbooks.Add() }
    $sheet=$book.Worksheets.Item(1)
    if (-not $WorkbookPath) {
        $sheet.Name='样式对照'; $sheet.Columns.Item(1).ColumnWidth=32; $sheet.Rows.RowHeight=45
        $named=$book.Styles.Add('QuoteMoney'); $named.NumberFormat='"¥"#,##0.000'; $named.Font.Bold=$true
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($named)
        $values=@('默认中文 ABC','标题中文 ABC',1234.5,'旋转 45°','负向旋转','缩进','Shrink this long amount description','Named style')
        for($i=0;$i -lt $values.Count;$i++) {
            $cell=$sheet.Cells.Item($i+1,1)
            try {
                if ($values[$i] -is [string]) { $cell.Value2=[string]$values[$i] } else {$cell.Value2=[double]$values[$i]}
                switch ($i) {
                    0 {$cell.Font.ThemeFont=2}
                    1 {$cell.Font.ThemeFont=1}
                    2 {$cell.Style='QuoteMoney'}
                    3 {$cell.Orientation=45}
                    4 {$cell.Orientation=-45}
                    5 {$cell.IndentLevel=2}
                    6 {$cell.ShrinkToFit=$true}
                    7 {$cell.Style='QuoteMoney'}
                }
            } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)}
        }
    }
    $rows=@()
    for($row=1;$row -le 8;$row++) {
        $cell=$sheet.Cells.Item($row,1); $font=$cell.Font; $cellStyle=$cell.Style
        try {
            $rows += [ordered]@{Row=$row;Value=$cell.Value2;Text=$cell.Text;Style=[string]$cellStyle.Name;NumberFormat=$cell.NumberFormat;
                Font=$font.Name;Size=$font.Size;Bold=$font.Bold;Italic=$font.Italic;ThemeFont=$font.ThemeFont;
                Orientation=$cell.Orientation;Indent=$cell.IndentLevel;Shrink=$cell.ShrinkToFit}
        } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cellStyle);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($font);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)}
    }
    if (-not $WorkbookPath) {$book.SaveAs((Join-Path $target 'excel-source.xlsx'),51)}
    [ordered]@{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=$excel.Calculation;
        Date1904=$book.Date1904;CorruptLoad='xlNormalLoad';Cases=$rows} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $target 'excel-styles.json') -Encoding utf8
    Write-Output $target
} finally {
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($sheet){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
