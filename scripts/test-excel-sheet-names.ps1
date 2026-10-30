param([Parameter(Mandatory)][string]$OutputDirectory, [string]$VerifyDirectory, [string]$CaseId)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Hash-File([string]$path){$stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);try{return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream))}finally{$stream.Dispose()}}
function Capture($book){
    $sheets=@();$names=@()
    for($i=1;$i -le $book.Names.Count;$i++){$name=$book.Names.Item($i);try{if($name.Visible){$names+=@{Name=[string]$name.Name;Formula=[string]$name.RefersTo;Visible=[bool]$name.Visible}}}finally{Release-Com $name}}
    for($s=1;$s -le $book.Worksheets.Count;$s++){
        $sheet=$book.Worksheets.Item($s);$cells=@();$shapes=@()
        try{
            for($r=1;$r -le 6;$r++){for($c=1;$c -le 6;$c++){
                $cell=$sheet.Cells.Item($r,$c);$font=$cell.Font
                try{
                    $value=$cell.Value2;$kind='Blank';if($null -ne $value){if($value -is [string]){$kind='Text'}elseif($value -is [bool]){$kind='Boolean'}elseif([string]$cell.Text -match '^#'){$kind='Error';$value=[string]$cell.Text}else{$kind='Number'}}
                    $cells+=@{Row=$r-1;Column=$c-1;Kind=$kind;Value=$value;Formula=$(if($cell.HasFormula){[string]$cell.Formula}else{$null});Format=[string]$cell.NumberFormat;Bold=[bool]$font.Bold;Font=[string]$font.Name;Size=[double]$font.Size}
                }finally{Release-Com $font;Release-Com $cell}
            }}
            for($i=1;$i -le $sheet.Shapes.Count;$i++){
                $shape=$sheet.Shapes.Item($i);$tl=$shape.TopLeftCell;$br=$shape.BottomRightCell
                try{$shapes+=@{Type=[int]$shape.Type;TopLeft=$tl.Address($false,$false);BottomRight=$br.Address($false,$false);Left=[double]$shape.Left;Top=[double]$shape.Top;Width=[double]$shape.Width;Height=[double]$shape.Height}}
                finally{Release-Com $br;Release-Com $tl;Release-Com $shape}
            }
            $sheets+=@{Name=[string]$sheet.Name;Visible=[int]$sheet.Visible;Cells=$cells;Shapes=$shapes}
        }finally{Release-Com $sheet}
    }
    return @{Sheets=$sheets;Names=@($names|Sort-Object {$_.Name})}
}
$ids=@('copy','reorder','rename-global','rename-local','edit-global','delete-global','delete-local','create-shadow')
if($CaseId){if($CaseId -notin $ids){throw 'Unknown case.'};$ids=@($CaseId)}
$excel=$null;$book=$null;$results=@()
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $pixel=Join-Path $target 'pixel.png';[IO.File]::WriteAllBytes($pixel,[Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aA1sAAAAASUVORK5CYII='))
    foreach($id in $ids){
        if($VerifyDirectory -and $id -eq 'delete-local'){continue} # The product explicitly rejects qualified local-name deletion.
        if($VerifyDirectory){$path=Join-Path ([IO.Path]::GetFullPath($VerifyDirectory,$repo)) "$id.xlsx";$hash=Hash-File $path;$book=$excel.Workbooks.Open($path,0,$true)}else{
            $book=$excel.Workbooks.Add();while($book.Worksheets.Count -gt 1){$remove=$book.Worksheets.Item(2);$remove.Delete();Release-Com $remove}
            $source=$book.Worksheets.Item(1);$source.Name='Source'
            $other=$book.Worksheets.Add([Type]::Missing,$source);$other.Name='Other'
            $watch=$book.Worksheets.Add([Type]::Missing,$other);$watch.Name='Watch'
            try{
                foreach($sheet in @($source,$other,$watch)){
                    $range=$sheet.Range('A1:F6');$font=$range.Font
                    try{$font.Name='Arial';$font.Size=11;$font.Bold=$false;$range.ColumnWidth=14;$range.RowHeight=24}finally{Release-Com $font;Release-Com $range}
                }
                foreach($r in @(1,2,3)){$cell=$source.Cells.Item($r,1);try{$cell.Value2=[double]($r*10)}finally{Release-Com $cell}}
                $cell=$other.Range('A1');try{$cell.Value2=[double]100}finally{Release-Com $cell}
                foreach($definition in @(@('GlobalSource','=Source!$A$1'),@('GlobalOther','=Other!$A$1'),@('Rate','=Source!$A$1'),@('Source!Rate','=Source!$A$2'),
                    @('Source!Local','=Source!$A$3'),@('Chain','=GlobalSource*2'),@('Source!LocalChain','=Source!Local+Source!Rate'),@('Other!Local','=Other!$A$1'),
                    @('GlobalRateAlias','=Rate'),@('UnusedSource','=Source!$A$3'),@('UnusedOther','=Other!$A$1'))){
                    $name=$book.Names.Add($definition[0],$definition[1]);Release-Com $name
                }
                foreach($pair in @(@('B1','=A1+Source!$A$2'),@('B2','=GlobalSource'),@('B3','=Rate'),@('B4','=Local'),@('B5','=Source!Local'),@('B6','=GlobalOther'),
                    @('C1','=Chain'),@('C2','=LocalChain'),@('C3','=Other!Local'),@('C4','="Source!A1"'),@('C5','=GlobalRateAlias'))){$cell=$source.Range($pair[0]);try{$cell.Formula=$pair[1]}finally{Release-Com $cell}}
                foreach($pair in @(@('B1','=GlobalSource+Rate'),@('B2','=Source!Rate'))){$cell=$other.Range($pair[0]);try{$cell.Formula=$pair[1]}finally{Release-Com $cell}}
                foreach($pair in @(@('B1','=Source!Local'),@('B2','=Rate'))){$cell=$watch.Range($pair[0]);try{$cell.Formula=$pair[1]}finally{Release-Com $cell}}
                $cell=$source.Range('A1');$font=$cell.Font;try{$font.Bold=$true;$cell.NumberFormat='0.00'}finally{Release-Com $font;Release-Com $cell}
                $range=$source.Range('D4:E5');$shapes=$source.Shapes
                try{$shape=$shapes.AddPicture($pixel,0,-1,$range.Left,$range.Top,$range.Width,$range.Height);Release-Com $shape}finally{Release-Com $shapes;Release-Com $range}
                $excel.CalculateFullRebuild();$path=Join-Path $target "$id-source.xlsx";$book.SaveAs($path,51);$hash=Hash-File $path
                switch($id){
                    'copy' {$source.Copy([Type]::Missing,$other);$copy=$book.Worksheets.Item(3);try{$copy.Name='Copy'}finally{Release-Com $copy}}
                    'reorder' {$source.Move([Type]::Missing,$watch)}
                    'rename-global' {$name=$book.Names.Item('GlobalSource');try{$name.Name='Amount'}finally{Release-Com $name}}
                    'rename-local' {$name=$book.Names.Item('Source!Rate');try{$name.Name='LocalRate'}finally{Release-Com $name}}
                    'edit-global' {$name=$book.Names.Item('GlobalSource');try{$name.RefersTo='=Other!$A$1'}finally{Release-Com $name}}
                    'delete-local' {$name=$book.Names.Item('Source!Rate');try{$name.Delete()}finally{Release-Com $name}}
                    'delete-global' {$name=$book.Names.Item('GlobalOther');try{$name.Delete()}finally{Release-Com $name}}
                    'create-shadow' {$name=$book.Names.Add('Source!GlobalOther','=Source!$A$3');Release-Com $name}
                }
            }finally{Release-Com $watch;Release-Com $other;Release-Com $source}
        }
        $excel.CalculateFullRebuild();$initial=Capture $book;$first=Join-Path $target "$id-first.xlsx";$book.SaveAs($first,51)
        $book.Close($false);Release-Com $book;$book=$null;$book=$excel.Workbooks.Open($first,0,$true)
        $excel.CalculateFullRebuild();$before=Capture $book;$output=Join-Path $target "$id.xlsx";$book.SaveAs($output,51)
        $book.Close($false);Release-Com $book;$book=$null;$book=$excel.Workbooks.Open($output,0,$true);$excel.CalculateFullRebuild();$after=Capture $book
        $results+=@{Id=$id;ExpectedGuard=($id -eq 'delete-local');SourceSha256=$hash;SourceUnchanged=((Hash-File $path) -eq $hash);BeforeInitialSave=$initial;BeforeSave=$before;Reopened=$after;
            Stable=(($before|ConvertTo-Json -Depth 10 -Compress) -ceq ($after|ConvertTo-Json -Depth 10 -Compress))}
        $book.Close($false);Release-Com $book;$book=$null;Write-Output "$id captured and reopened."
    }
    @{Origin=$(if($VerifyDirectory){'excel-output-verification'}else{'independent-excel-oracle'});Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Cases=$results}|
        ConvertTo-Json -Depth 14|Set-Content (Join-Path $target 'excel-sheet-names.json') -Encoding utf8
    if(@($results|Where-Object {(-not $_.Stable -and -not $_.ExpectedGuard) -or -not $_.SourceUnchanged}).Count){throw 'Sheet/name Excel save/reopen mismatch.'}
}finally{
    try{if($book){$book.Close($false);Release-Com $book}}
    finally{[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
}
