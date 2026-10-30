param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$excel=$null;$book=$null;$cases=@()
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    foreach($axis in @('row','column')){foreach($index in @(1,2)){
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$range=$sheet.Range('A1:C4');$font=$range.Font
        try{
            $blank=$sheet.Range('D5');$defaultFont=$blank.Font
            try{$defaults=@{Font=[string]$defaultFont.Name;Size=[double]$defaultFont.Size;Bold=[bool]$defaultFont.Bold;RowHeight=[double]$blank.RowHeight;ColumnWidth=[double]$blank.ColumnWidth}}
            finally{Release-Com $defaultFont;Release-Com $blank}
            $font.Name='Arial';$font.Size=14;$font.Bold=$true;$range.RowHeight=27;$range.ColumnWidth=22
            if($axis -eq 'row'){$neighbor=$sheet.Rows.Item(1)}else{$neighbor=$sheet.Columns.Item(1)}
            try{$neighbor.Hidden=$true}finally{Release-Com $neighbor}
            if($axis -eq 'row'){$insert=$sheet.Range("${index}:$($index+1)")}else{$insert=$sheet.Range($(if($index -eq 1){'A:B'}else{'B:C'}))}
            try{[void]$insert.Insert()}finally{Release-Com $insert}
            if($axis -eq 'row'){$cell=$sheet.Cells.Item($index,1);$newAxis=$sheet.Rows.Item($index)}else{$cell=$sheet.Cells.Item(1,$index);$newAxis=$sheet.Columns.Item($index)}
            $f=$cell.Font
            try{
                $expected=[ordered]@{Font=$(if($index -eq 1){$defaults.Font}else{'Arial'});Size=$(if($index -eq 1){$defaults.Size}else{[double]14});
                    Bold=$(if($index -eq 1){$defaults.Bold}else{$true});RowHeight=$(if($axis -eq 'row' -and $index -eq 1){$defaults.RowHeight}else{[double]27});
                    ColumnWidth=$(if($axis -eq 'column' -and $index -eq 1){$defaults.ColumnWidth}else{[double]22});Hidden=$false}
                $actual=[ordered]@{Font=[string]$f.Name;Size=[double]$f.Size;Bold=[bool]$f.Bold;RowHeight=[double]$cell.RowHeight;ColumnWidth=[double]$cell.ColumnWidth;Hidden=[bool]$newAxis.Hidden}
                $cases += [ordered]@{Axis=$axis;Index=$index;Count=2;Expected=$expected;Actual=$actual;Passed=(($expected|ConvertTo-Json -Compress) -ceq ($actual|ConvertTo-Json -Compress))}
            }finally{Release-Com $newAxis;Release-Com $f;Release-Com $cell}
        }finally{Release-Com $font;Release-Com $range;Release-Com $sheet;$book.Close($false);Release-Com $book;$book=$null}
    }}
    [ordered]@{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Cases=$cases}|ConvertTo-Json -Depth 8|
        Set-Content -LiteralPath (Join-Path $target 'excel-insert-formatting.json') -Encoding utf8
    if(@($cases|Where-Object { -not $_.Passed }).Count){throw 'Excel insertion boundary changed; see evidence.'}
    Write-Output 'Four Excel insertion boundary cases passed.'
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
}
