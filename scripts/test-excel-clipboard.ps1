param([Parameter(Mandatory)][string]$OutputDirectory, [string]$VerifyDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Hash-File([string]$path){$stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);try{return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream))}finally{$stream.Dispose()}}
function Capture($book){
    $sheets=@()
    for($s=1;$s -le 2;$s++){
        $sheet=$book.Worksheets.Item($s);$cells=@()
        try{for($r=1;$r -le 10;$r++){for($c=1;$c -le 8;$c++){
            $cell=$sheet.Cells.Item($r,$c);$font=$cell.Font
            try{
                $value=$cell.Value2;$kind='Blank';$formula=$null
                if($cell.HasFormula){$formula=[string]$cell.Formula}
                if($null -ne $value){if($value -is [string]){$kind='Text'}elseif($value -is [bool]){$kind='Boolean'}elseif([string]$cell.Text -match '^#'){$kind='Error';$value=[string]$cell.Text}else{$kind='Number'}}
                $chars=@()
                if($kind -eq 'Text') {for($i=1;$i -le $value.Length;$i++){
                    $ch=$cell.Characters($i,1);$f=$ch.Font
                    try{$chars+=@{Name=[string]$f.Name;Size=[double]$f.Size;Bold=[bool]$f.Bold;Italic=[bool]$f.Italic;Color=[int]$f.Color}}
                    finally{Release-Com $f;Release-Com $ch}
                }}
                $cells += [ordered]@{Row=$r-1;Column=$c-1;Kind=$kind;Value=$value;Formula=$formula;NumberFormat=[string]$cell.NumberFormat;
                    Font=[string]$font.Name;Size=$font.Size;Bold=$font.Bold;Italic=$font.Italic;Color=$font.Color;Characters=$chars}
            }finally{Release-Com $font;Release-Com $cell}
        }}}finally{Release-Com $sheet}
        $sheets+=@{Index=$s-1;Cells=$cells}
    }
    return $sheets
}
$excel=$null;$book=$null;$cases=@()
$modes=@('all','all-sheet','values','formulas','formats','transpose','transpose-mixed','transpose-qualified','cut-overlap','cut-sheet','rich-all','rich-values','rich-formulas','rich-formats')
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    foreach($mode in $modes){
        if($VerifyDirectory -and $mode -in @('rich-values','rich-formulas','rich-formats')){continue}
        if($VerifyDirectory){$path=Join-Path ([IO.Path]::GetFullPath($VerifyDirectory,$repo)) "$mode.xlsx";$hash=Hash-File $path;$book=$excel.Workbooks.Open($path,0,$true)}else{
            $book=$excel.Workbooks.Add();while($book.Worksheets.Count -gt 1){$remove=$book.Worksheets.Item(2);$remove.Delete();Release-Com $remove}
            $sheet=$book.Worksheets.Item(1);$sheet.Name='Source';$second=$book.Worksheets.Add();$second.Name='Target';$second.Move([Type]::Missing,$sheet)
            foreach($ws in @($sheet,$second)){
                $range=$ws.Range('A1:H10');$font=$range.Font
                try{$font.Name='Arial';$font.Size=11;$font.Bold=$false;$font.Italic=$false;$font.Color=0;$range.ColumnWidth=24}
                finally{Release-Com $font;Release-Com $range}
            }
            try{
                foreach($entry in @(@('A1','00123'),@('A2','source'),@('E5','old'),@('F5','old'),@('E6','old'))){
                    $cell=$sheet.Range($entry[0]);try{$cell.NumberFormat='@';$cell.Value2=[string]$entry[1]}finally{Release-Com $cell}
                }
                foreach($entry in @(@('B1',12.5),@('A3',10),@('B3',20),@('A4',30),@('B4',40))){
                    $cell=$sheet.Range($entry[0]);try{$cell.Value2=[double]$entry[1]}finally{Release-Com $cell}
                }
                $cell=$sheet.Range('A1');$font=$cell.Font;try{$font.Bold=$true;$font.Size=12}finally{Release-Com $font;Release-Com $cell}
                $cell=$sheet.Range('B1');$font=$cell.Font;try{$cell.NumberFormat='0.00';$font.Italic=$true}finally{Release-Com $font;Release-Com $cell}
                $cell=$sheet.Range('B2');$font=$cell.Font;try{$font.Color=255}finally{Release-Com $font;Release-Com $cell}
                $cell=$sheet.Range('C1');try{$cell.Formula='=A3+$A$3+$A3+A$3+SUM(A3:B4)'}finally{Release-Com $cell}
                if($mode -eq 'transpose-mixed'){$cell=$sheet.Range('C1');try{$cell.Formula='=SUM($A3:B4,A$3:B4,A3:$B4,A3:B$4)'}finally{Release-Com $cell}}
                if($mode -eq 'transpose-qualified'){$cell=$sheet.Range('C1');try{$cell.Formula='=SUM(Target!A3:B4)+Target!A3+Target!$A3+Target!A$3'}finally{Release-Com $cell}}
                $cell=$sheet.Range('C2');try{$cell.Formula='=1/0'}finally{Release-Com $cell}
                $cell=$second.Range('H10');try{$cell.Formula='=Source!A1'}finally{Release-Com $cell}
                if($mode.StartsWith('rich-')){
                    $cell=$sheet.Range('A1');$ch=$cell.Characters(2,2);$font=$ch.Font
                    try{$font.Italic=$true;$font.Color=255}finally{Release-Com $font;Release-Com $ch;Release-Com $cell}
                    $cell=$sheet.Range('E5');$ch=$cell.Characters(1,1);$font=$ch.Font
                    try{$font.Italic=$true;$font.Size=16}finally{Release-Com $font;Release-Com $ch;Release-Com $cell}
                }
                $excel.CalculateFullRebuild()
                $book.SaveAs((Join-Path $target "$mode-source.xlsx"),51)
                $path=Join-Path $target "$mode-source.xlsx";$hash=Hash-File $path
                $source=$sheet.Range('A1:C2');$dest=$sheet.Range('E5')
                if($mode -eq 'cut-overlap'){Release-Com $dest;$dest=$sheet.Range('B2')}
                if($mode -in @('cut-sheet','all-sheet')){Release-Com $dest;$dest=$second.Range('E5')}
                try{
                    if($mode.StartsWith('cut-')){[void]$source.Cut($dest)}
                    elseif($mode -in @('all','all-sheet','rich-all')){[void]$source.Copy($dest)}
                    else{
                        [void]$source.Copy()
                        $paste=-4104
                        if($mode -match 'values$'){$paste=-4163}elseif($mode -match 'formulas$'){$paste=-4123}elseif($mode -match 'formats$'){$paste=-4122}
                        [void]$dest.PasteSpecial($paste,-4142,$false,$mode.StartsWith('transpose'))
                    }
                }finally{Release-Com $dest;Release-Com $source}
                $excel.CutCopyMode=0
            }finally{Release-Com $sheet;Release-Com $second}
        }
        $excel.CalculateFullRebuild()
        $before=@(Capture $book)
        $output=Join-Path $target "$mode.xlsx"
        $book.SaveAs($output,51);$book.Close($false);Release-Com $book;$book=$null
        $book=$excel.Workbooks.Open($output,0,$true);$after=@(Capture $book)
        $cases+=@{Mode=$mode;SourceSha256=$hash;SourceUnchanged=((Hash-File $path) -eq $hash);Sheets=$before;Reopened=$after;Stable=(($before|ConvertTo-Json -Depth 12 -Compress) -ceq ($after|ConvertTo-Json -Depth 12 -Compress))}
        $book.Close($false);Release-Com $book;$book=$null
        Write-Output "$mode captured and reopened."
    }
    @{Origin=$(if($VerifyDirectory){'excel-output-verification'}else{'independent-excel-oracle'});Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Cases=$cases}|ConvertTo-Json -Depth 16|
        Set-Content -LiteralPath (Join-Path $target 'excel-clipboard.json') -Encoding utf8
    if(@($cases|Where-Object {-not $_.Stable}).Count){throw 'Excel save/reopen mismatch.'}
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
}
