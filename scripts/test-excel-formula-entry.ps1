param([Parameter(Mandatory)][string]$OutputDirectory, [Parameter(Mandatory)][string]$QuickExcelPath)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Capture($book){
    $sheet=$book.Worksheets.Item(1);$values=@()
    try{for($r=1;$r -le 5;$r++){for($c=1;$c -le 3;$c++){
        $cell=$sheet.Cells.Item($r,$c)
        try{$values+=@{Row=$r;Column=$c;Value=$cell.Value2;Formula=$(if($cell.HasFormula){[string]$cell.Formula}else{$null})}}
        finally{Release-Com $cell}
    }}}finally{Release-Com $sheet}
    return $values
}
$excel=$null;$book=$null;$navigation=@();$conversions=@();$actualPath=[IO.Path]::GetFullPath($QuickExcelPath,$repo)
$sourceHash=(Get-FileHash -LiteralPath $actualPath -Algorithm SHA256).Hash
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
    try{
        $sheet.Name='Entry'
        foreach($row in @(1,2,3,5)){$cell=$sheet.Cells.Item($row,1);try{$cell.Value2=[double]$row}finally{Release-Com $cell}}
        foreach($hidden in @($false,$true)){
            $row=$sheet.Rows.Item(2);try{$row.Hidden=$hidden}finally{Release-Com $row}
            foreach($case in @(@{From='A1';Direction=-4121},@{From='A3';Direction=-4121},@{From='A4';Direction=-4121},@{From='A5';Direction=-4121},@{From='A5';Direction=-4162},@{From='A3';Direction=-4162})){
                $cell=$sheet.Range($case.From);$end=$cell.End($case.Direction)
                try{$navigation+=@{From=$case.From;Direction=$case.Direction;HiddenRow2=$hidden;To=$end.Address($false,$false)}}finally{Release-Com $end;Release-Com $cell}
            }
        }
        $row=$sheet.Rows.Item(2);try{$row.Hidden=$false}finally{Release-Com $row}
        $cell=$sheet.Range('A3');try{$cell.Formula='=""'}finally{Release-Com $cell}
        $cell=$sheet.Range('A1');$end=$cell.End(-4121);try{$navigation+=@{From='A1';Direction=-4121;FormulaEmptyA3=$true;To=$end.Address($false,$false)}}finally{Release-Com $end;Release-Com $cell}
        $cell=$sheet.Range('A3');try{$cell.Value2=[double]3}finally{Release-Com $cell}
        foreach($mode in @(1,2,3,4)){$conversions+=@{Mode=$mode;Formula=$excel.ConvertFormula('=A1+B2',1,1,$mode)}}
        foreach($pair in @(@('C1','=SUM(A1:A3)'),@('C2','=SUM(A1:A2)'),@('C3','=$A$1+$A2'))){
            $cell=$sheet.Range($pair[0]);try{$cell.Formula=$pair[1]}finally{Release-Com $cell}
        }
        $excel.CalculateFullRebuild();$expected=Capture $book;$oracle=Join-Path $target 'excel-entry-oracle.xlsx';$book.SaveAs($oracle,51)
    }finally{Release-Com $sheet}
    $book.Close($false);Release-Com $book;$book=$excel.Workbooks.Open($oracle,0,$true);$expectedReopened=Capture $book
    $book.Close($false);Release-Com $book;$book=$excel.Workbooks.Open($actualPath,0,$true)
    $excel.CalculateFullRebuild();$actual=Capture $book;$book.SaveAs((Join-Path $target 'excel-resaved.xlsx'),51)
    $book.Close($false);Release-Com $book;$book=$excel.Workbooks.Open((Join-Path $target 'excel-resaved.xlsx'),0,$true);$reopened=Capture $book
    $stable=(($expected|ConvertTo-Json -Depth 6 -Compress) -ceq ($expectedReopened|ConvertTo-Json -Depth 6 -Compress))
    $matches=(($expected|ConvertTo-Json -Depth 6 -Compress) -ceq ($actual|ConvertTo-Json -Depth 6 -Compress))
    $actualStable=(($actual|ConvertTo-Json -Depth 6 -Compress) -ceq ($reopened|ConvertTo-Json -Depth 6 -Compress))
    $book.Close($false);Release-Com $book;$book=$null
    $unchanged=(Get-FileHash -LiteralPath $actualPath -Algorithm SHA256).Hash -eq $sourceHash
    @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Origin='independent Excel Range.End/ConvertFormula and formula oracle';
        Navigation=$navigation;ReferenceModes=$conversions;Expected=$expected;Actual=$actual;Reopened=$reopened;OracleStable=$stable;Matches=$matches;ActualStable=$actualStable;
        SourceSha256=$sourceHash;SourceUnchanged=$unchanged;Passed=($stable -and $matches -and $actualStable -and $unchanged)}|
        ConvertTo-Json -Depth 10|Set-Content (Join-Path $target 'excel-formula-entry.json') -Encoding utf8
    if(-not ($stable -and $matches -and $actualStable -and $unchanged)){throw 'Formula entry Excel comparison failed.'}
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}
}
