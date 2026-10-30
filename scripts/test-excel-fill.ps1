param([Parameter(Mandatory)][string]$OutputDirectory, [string]$VerifyDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Hash-File([string]$path){$stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);try{return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream))}finally{$stream.Dispose()}}
function Capture($book){
    $sheet=$book.Worksheets.Item(1);$cells=@()
    try{for($r=1;$r -le 12;$r++){for($c=1;$c -le 6;$c++){
        $cell=$sheet.Cells.Item($r,$c);$font=$cell.Font
        try{
            $value=$cell.Value2;$kind='Blank';$formula=$null
            if($cell.HasFormula){$formula=[string]$cell.Formula}
            if($null -ne $value){if($value -is [string]){$kind='Text'}elseif($value -is [bool]){$kind='Boolean'}elseif([string]$cell.Text -match '^#'){$kind='Error';$value=[string]$cell.Text}else{$kind='Number'}}
            $characters=@()
            if($kind -eq 'Text'){for($i=1;$i -le $value.Length;$i++){
                $character=$cell.Characters($i,1);$characterFont=$character.Font
                try{$characters+=@{Name=[string]$characterFont.Name;Size=[double]$characterFont.Size;Bold=[bool]$characterFont.Bold;Italic=[bool]$characterFont.Italic;Color=[int]$characterFont.Color}}
                finally{Release-Com $characterFont;Release-Com $character}
            }}
            $cells+=@{Row=$r-1;Column=$c-1;Kind=$kind;Value=$value;Formula=$formula;Format=[string]$cell.NumberFormat;Text=[string]$cell.Text;
                Font=[string]$font.Name;Size=$font.Size;Bold=$font.Bold;Italic=$font.Italic;Color=$font.Color;Characters=$characters}
        }finally{Release-Com $font;Release-Com $cell}
    }}}finally{Release-Com $sheet}
    return @{Date1904=[bool]$book.Date1904;Cells=$cells}
}
$definitions=@(
    @{Id='down';Operation='down'},@{Id='right';Operation='right'},@{Id='handle-down';Operation='handle-down'},@{Id='handle-left';Operation='handle-left'},
    @{Id='clear-content';Operation='clear-content'},@{Id='clear-format';Operation='clear-format'},
    @{Id='numbers';Operation='series';Kind='Numbers';Start=0.1;Step=0.1},@{Id='numbers-negative';Operation='series';Kind='Numbers';Start=3;Step=-2},
    @{Id='numbers-right';Operation='series-right';Kind='Numbers';Start=3;Step=2},
    @{Id='days-1900';Operation='series';Kind='Days';Start=59;Step=1},@{Id='days-1904';Operation='series';Kind='Days';Start=0;Step=1;Date1904=$true},
    @{Id='months-end';Operation='series';Kind='Months';Start=31;Step=1},@{Id='months-fictitious';Operation='series';Kind='Months';Start=60;Step=1},
    @{Id='months-fraction';Operation='series';Kind='Months';Start=45322.5;Step=1},@{Id='months-back';Operation='series';Kind='Months';Start=45322;Step=-1},
    @{Id='years-leap';Operation='series';Kind='Years';Start=45351;Step=1},@{Id='years-1904';Operation='series';Kind='Years';Start=59;Step=1;Date1904=$true})
$excel=$null;$book=$null;$results=@()
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    foreach($definition in $definitions){
        $id=$definition.Id
        if($VerifyDirectory){$path=Join-Path ([IO.Path]::GetFullPath($VerifyDirectory,$repo)) "$id.xlsx";$hash=Hash-File $path;$book=$excel.Workbooks.Open($path,0,$true)}else{
            $book=$excel.Workbooks.Add();while($book.Worksheets.Count -gt 1){$remove=$book.Worksheets.Item(2);$remove.Delete();Release-Com $remove}
            $book.Date1904=($definition.Date1904 -eq $true);$sheet=$book.Worksheets.Item(1);$sheet.Name='Fill'
            try{
                $range=$sheet.Range('A1:F12');$font=$range.Font
                try{$range.ColumnWidth=24;$font.Name='Arial';$font.Size=11;$font.Bold=$false;$font.Italic=$false;$font.Color=0}finally{Release-Com $font;Release-Com $range}
                $cell=$sheet.Range('A1');$font=$cell.Font
                try{$cell.Value2=[double]1;$font.Bold=$true;$font.Size=14}finally{Release-Com $font;Release-Com $cell}
                if($definition.Operation.StartsWith('series')){
                    $cell=$sheet.Range('A1');try{$cell.Value2=[double]$definition.Start;if($definition.Kind -ne 'Numbers'){$cell.NumberFormat='yyyy-mm-dd hh:mm'}}finally{Release-Com $cell}
                }else{
                    $cell=$sheet.Range('A2');try{$cell.Value2=[double]2}finally{Release-Com $cell}
                    $cell=$sheet.Range('F1');try{$cell.Value2=[double]10}finally{Release-Com $cell}
                    $cell=$sheet.Range('B1');try{$cell.Formula='=A1+$F$1'}finally{Release-Com $cell}
                    $cell=$sheet.Range('B2');try{$cell.Formula='=A2+$F$1'}finally{Release-Com $cell}
                    if($definition.Operation -eq 'right'){$cell=$sheet.Range('A2');try{$cell.Formula='=A1+$F$1'}finally{Release-Com $cell}}
                    if($definition.Operation -in @('clear-content','clear-format')){
                        $cell=$sheet.Range('A1');try{$cell.Value2='plain ID';$cell.NumberFormat='@'}finally{Release-Com $cell}
                        $cell=$sheet.Range('A1');$characters=$cell.Characters(2,2);$font=$characters.Font
                        try{$font.Italic=$true;$font.Color=255}finally{Release-Com $font;Release-Com $characters;Release-Com $cell}
                        $cell=$sheet.Range('B1');try{$cell.Formula='=A2+1'}finally{Release-Com $cell}
                    }
                    if($definition.Operation -eq 'handle-left'){
                        $a=$sheet.Range('A1:B2');$b=$sheet.Range('C1');try{[void]$a.Copy($b)}finally{Release-Com $b;Release-Com $a}
                    }
                }
                $excel.CalculateFullRebuild();$path=Join-Path $target "$id-source.xlsx";$book.SaveAs($path,51);$hash=Hash-File $path
                $range=$sheet.Range('A1:B6');$source=$null
                try{switch($definition.Operation){
                    'down' {[void]$range.FillDown()}
                    'right' {Release-Com $range;$range=$sheet.Range('A1:E2');[void]$range.FillRight()}
                    'handle-down' {$source=$sheet.Range('A1:B2');[void]$source.AutoFill($range,1)}
                    'handle-left' {Release-Com $range;$range=$sheet.Range('A1:D2');$source=$sheet.Range('C1:D2');[void]$source.AutoFill($range,1)}
                    'clear-content' {[void]$range.ClearContents()}
                    'clear-format' {[void]$range.ClearFormats()}
                    default {
                        Release-Com $range;$range=$sheet.Range($(if($definition.Operation -eq 'series-right'){'A1:F1'}else{'A1:A8'}))
                        $type=-4132;$unit=1
                        if($definition.Kind -ne 'Numbers'){$type=3;$unit=switch($definition.Kind){'Days'{1}'Months'{3}'Years'{4}}}
                        [void]$range.DataSeries($(if($definition.Operation -eq 'series-right'){1}else{2}),$type,$unit,[double]$definition.Step,[Type]::Missing,$false)
                    }
                }}finally{Release-Com $source;Release-Com $range}
                $excel.CutCopyMode=0
            }finally{Release-Com $sheet}
        }
        $excel.CalculateFullRebuild();$before=Capture $book;$output=Join-Path $target "$id.xlsx";$book.SaveAs($output,51)
        $book.Close($false);Release-Com $book;$book=$null;$book=$excel.Workbooks.Open($output,0,$true);$after=Capture $book
        $results+=@{Id=$id;Definition=$definition;SourceSha256=$hash;SourceUnchanged=((Hash-File $path) -eq $hash);BeforeSave=$before;Reopened=$after;
            Stable=(($before|ConvertTo-Json -Depth 8 -Compress) -ceq ($after|ConvertTo-Json -Depth 8 -Compress))}
        $book.Close($false);Release-Com $book;$book=$null;Write-Output "$id captured and reopened."
    }
    @{Origin=$(if($VerifyDirectory){'excel-output-verification'}else{'independent-excel-oracle'});Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Cases=$results}|
        ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $target 'excel-fill.json') -Encoding utf8
    if(@($results|Where-Object {-not $_.Stable -or -not $_.SourceUnchanged}).Count){throw 'Excel save/reopen/source mismatch.'}
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}
}
