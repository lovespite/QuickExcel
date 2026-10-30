param([Parameter(Mandatory)][string]$OutputDirectory, [string]$InputDirectory, [string]$ReferencePath)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
if([bool]$InputDirectory -ne [bool]$ReferencePath){throw 'InputDirectory and ReferencePath must be supplied together.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($object){if($null -ne $object){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}}
# Separate call sites avoid PowerShell caching a numeric COM setter for text and vice versa.
function Set-Text($sheet,[string]$address,[string]$value){$cell=$sheet.Range($address);try{$cell.Value2=$value}finally{Release-Com $cell}}
function Set-Number($sheet,[string]$address,[double]$value){$cell=$sheet.Range($address);try{$cell.Value2=$value}finally{Release-Com $cell}}
function Set-Formula($sheet,[string]$address,[string]$value){$cell=$sheet.Range($address);try{$cell.Formula=$value}finally{Release-Com $cell}}
function Set-Format($sheet,[string]$address,[string]$value){
    # IDispatch uses the host LCID, including the nominally invariant NumberFormat property.
    if([Globalization.CultureInfo]::CurrentCulture.TwoLetterISOLanguageName -eq 'zh'){$value=$value.Replace('[Red]','[红色]')}
    $cell=$sheet.Range($address);try{$cell.NumberFormat=$value}finally{Release-Com $cell}
}
function Capture-Book($book){
    $sheets=@()
    for($i=1;$i -le $book.Worksheets.Count;$i++){
        $sheet=$book.Worksheets.Item($i);$cells=@();$setup=$sheet.PageSetup
        try{
            # Fixed small bounds include blank cells and catch extra values; these are test fixtures only.
            for($row=1;$row -le 6;$row++){for($col=1;$col -le 4;$col++){
                $cell=$sheet.Cells.Item($row,$col);$font=$cell.Font
                try{
                    $value=$cell.Value2
                    $kind=if($null -eq $value){'Blank'}elseif($value -is [bool]){'Boolean'}elseif($value -is [string]){'Text'}else{'Number'}
                    $cells += [ordered]@{Row=$row;Column=$col;Kind=$kind;Value=$value;Formula=$(if($cell.HasFormula){[string]$cell.Formula}else{$null});
                        Text=[string]$cell.Text;NumberFormat=[string]$cell.NumberFormat;Font=[string]$font.Name;Size=[double]$font.Size;Bold=[bool]$font.Bold;Italic=[bool]$font.Italic;Color=[int]$font.Color}
                }finally{Release-Com $font;Release-Com $cell}
            }}
            $printing=[ordered]@{}
            foreach($name in @('PaperSize','Orientation','Zoom','FitToPagesWide','FitToPagesTall','PrintArea','PrintTitleRows','PrintTitleColumns','CenterHeader','RightFooter')){$printing[$name]=$setup.$name}
            $sheets += [ordered]@{Name=[string]$sheet.Name;Visible=[int]$sheet.Visible;Cells=$cells;Printing=$printing}
        }finally{Release-Com $setup;Release-Com $sheet}
    }
    return [ordered]@{Date1904=[bool]$book.Date1904;Sheets=$sheets}
}
function Initialize-Book($book,[string]$id){
    while($book.Worksheets.Count -gt 1){$extra=$book.Worksheets.Item($book.Worksheets.Count);try{$extra.Delete()}finally{Release-Com $extra}}
    $sheet=$book.Worksheets.Item(1)
    try{
        $sheet.Name='数据';$book.Date1904=($id -eq 'date-ledger-1904')
        Set-Text $sheet 'A1' '编号';Set-Text $sheet 'B1' '内容'
        switch($id){
            'roster' {
                Set-Format $sheet 'A2:A3' '@';Set-Text $sheet 'A2' '00123';Set-Text $sheet 'A3' '00007'
                Set-Text $sheet 'B2' '张三';Set-Text $sheet 'B3' '李四';Set-Text $sheet 'C3' "中文`n换行"
                $cell=$sheet.Range('C2');try{$cell.Value2=$true}finally{Release-Com $cell}
            }
            'quotation' {
                Set-Text $sheet 'A1' '单价';Set-Text $sheet 'B1' '数量';Set-Text $sheet 'C1' '金额'
                Set-Number $sheet 'A2' 1234.567;Set-Number $sheet 'B2' 2;Set-Formula $sheet 'C2' '=ROUND(A2*B2,2)'
                Set-Number $sheet 'A3' -42.5;Set-Number $sheet 'B3' 3;Set-Formula $sheet 'C3' '=A3*B3';Set-Formula $sheet 'C4' '=SUM(C2:C3)'
                Set-Format $sheet 'A2:A3' '"¥"#,##0.000;[Red]-"¥"#,##0.000';Set-Format $sheet 'C2:C4' '0.00'
            }
            {$_ -in 'date-ledger-1900','date-ledger-1904'} {
                Set-Text $sheet 'A1' '日期';Set-Text $sheet 'B1' '时间';Set-Number $sheet 'A2' 60;Set-Number $sheet 'A3' 45292.5
                Set-Number $sheet 'B2' 0.5;Set-Number $sheet 'B3' 1.5;Set-Format $sheet 'A2:A3' 'yyyy-mm-dd';Set-Format $sheet 'B2:B3' '[h]:mm:ss'
            }
            'cross-sheet' {
                Set-Text $sheet 'A1' '数量';Set-Number $sheet 'A2' 21;Set-Number $sheet 'B2' 3
                $other=$book.Worksheets.Add([Type]::Missing,$sheet)
                try{$other.Name='汇总';Set-Text $other 'A1' '跨表结果';Set-Formula $other 'A2' "='数据'!A2*2";Set-Formula $other 'B2' "=SUM('数据'!A2:B2)"}finally{Release-Com $other}
            }
            'print-template' {
                Set-Text $sheet 'A1' '打印报价';Set-Number $sheet 'A2' 12;Set-Number $sheet 'B2' 5;Set-Formula $sheet 'C2' '=A2*B2'
                $setup=$sheet.PageSetup
                try{$setup.PaperSize=9;$setup.Orientation=2;$setup.Zoom=$false;$setup.FitToPagesWide=1;$setup.FitToPagesTall=1;
                    $setup.PrintArea='$A$1:$D$6';$setup.PrintTitleRows='$1:$1';$setup.CenterHeader='业务基线';$setup.RightFooter='&P/&N'}finally{Release-Com $setup}
            }
        }
    }finally{Release-Com $sheet}
    for($i=1;$i -le $book.Worksheets.Count;$i++){
        $sheet=$book.Worksheets.Item($i);$range=$sheet.Range('A1:D6');$font=$range.Font;$header=$sheet.Range('A1:D1');$headerFont=$header.Font
        try{$font.Name='Arial';$font.Size=11;$font.Color=0;$headerFont.Bold=$true;$range.ColumnWidth=26;$range.RowHeight=24}
        finally{Release-Com $headerFont;Release-Com $header;Release-Com $font;Release-Com $range;Release-Com $sheet}
    }
}
$excel=$null;$book=$null;$results=@();$failures=@()
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $reference=if($ReferencePath){Get-Content -LiteralPath ([IO.Path]::GetFullPath($ReferencePath,$repo)) -Raw|ConvertFrom-Json}else{$null}
    foreach($id in @('roster','quotation','date-ledger-1900','date-ledger-1904','cross-sheet','print-template')){
        $inputPath=if($InputDirectory){Join-Path ([IO.Path]::GetFullPath($InputDirectory,$repo)) ($id+'.xlsx')}else{$null}
        $sourceHash=if($inputPath){(Get-FileHash -LiteralPath $inputPath).Hash}else{$null}
        if($inputPath){$book=$excel.Workbooks.Open($inputPath,0,$true)}else{$book=$excel.Workbooks.Add();Initialize-Book $book $id}
        $excel.CalculateFullRebuild();$calculation=[int]$excel.Calculation;$before=Capture-Book $book
        $saved=Join-Path $target ($id+'.xlsx');$book.SaveAs($saved,51);$book.Close($false);Release-Com $book;$book=$null
        $book=$excel.Workbooks.Open($saved,0,$true);$excel.CalculateFullRebuild();$after=Capture-Book $book
        $same=($before|ConvertTo-Json -Depth 12 -Compress) -ceq ($after|ConvertTo-Json -Depth 12 -Compress)
        $matches=if($reference){$expected=($reference.Cases|Where-Object Id -eq $id).AfterReopen;($expected|ConvertTo-Json -Depth 12 -Compress) -ceq ($after|ConvertTo-Json -Depth 12 -Compress)}else{$null}
        $unchanged= -not $inputPath -or $sourceHash -eq (Get-FileHash -LiteralPath $inputPath).Hash
        $results += [ordered]@{Id=$id;File=$id+'.xlsx';Sha256=(Get-FileHash -LiteralPath $saved).Hash;SourceSha256=$sourceHash;
            SourceUnchanged=$unchanged;SameAfterReopen=$same;ReferenceMatches=$matches;Before=$before;AfterReopen=$after}
        if(-not $same -or -not $unchanged -or $matches -eq $false){$failures+=$id}
        $book.Close($false);Release-Com $book;$book=$null
        Write-Output "$id Excel capture/reopen: $same"
    }
    [ordered]@{Origin='real-excel-generated-fixtures';Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);
        Calculation=$calculation;CorruptLoad='xlNormalLoad';ExternalLinksUpdated=$false;Cases=$results;Failures=$failures}|
        ConvertTo-Json -Depth 16|Set-Content -LiteralPath (Join-Path $target 'excel-business-baseline.json') -Encoding utf8
    if($failures.Count){throw "Excel comparison failed: $($failures -join ', ')"}
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
}
