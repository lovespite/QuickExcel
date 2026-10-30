param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath,[string]$ReferencePath,
    [ValidateSet('AfterReopen','AfterEdit')][string]$ReferenceState='AfterReopen')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($object){if($null -ne $object){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}}
function Set-Text($sheet,$address,[string]$text){$cell=$sheet.Range($address);try{$cell.Value2=$text}finally{Release-Com $cell}}
function Capture($book){
    $records=@()
    foreach($address in 'A1','A2','B1','C1','D1'){
        $sheet=$book.Worksheets.Item(1);$cell=$sheet.Range($address)
        try{
            $value=[string]$cell.Value2;$chars=@()
            for($i=1;$i -le $value.Length;$i++){
                $character=$cell.Characters($i,1);$font=$character.Font
                try{$chars += [ordered]@{Index=$i;Text=[string]$character.Text;Font=[string]$font.Name;Size=$font.Size;Color=$font.Color;Bold=$font.Bold;Italic=$font.Italic;Underline=$font.Underline}}
                finally{Release-Com $font;Release-Com $character}
            }
            $links=@();$collection=$cell.Hyperlinks
            try{for($i=1;$i -le $collection.Count;$i++){$link=$collection.Item($i);try{$links += [ordered]@{Address=$link.Address;SubAddress=$link.SubAddress;Text=$link.TextToDisplay;Tip=$link.ScreenTip}}finally{Release-Com $link}}}finally{Release-Com $collection}
            $comment=$cell.Comment;$note=$null
            if($comment){$shape=$comment.Shape;try{$note=[ordered]@{Author=$comment.Author;Text=$comment.Text();Visible=$comment.Visible;Width=[math]::Round($shape.Width,4);Height=[math]::Round($shape.Height,4);Left=[math]::Round($shape.Left,4);Top=[math]::Round($shape.Top,4)}}finally{Release-Com $shape;Release-Com $comment}}
            $records += [ordered]@{Cell=$address;Value=$value;Characters=$chars;Links=$links;Note=$note}
        }finally{Release-Com $cell;Release-Com $sheet}
    }
    return ,$records
}
function Equivalent($left,$right){
    $a=$left|ConvertTo-Json -Depth 25|ConvertFrom-Json;$b=$right|ConvertTo-Json -Depth 25|ConvertFrom-Json
    if($a.Count -ne $b.Count){return $false}
    for($i=0;$i -lt $a.Count;$i++){
        if(($null -eq $a[$i].Note) -ne ($null -eq $b[$i].Note)){return $false}
        if($a[$i].Note){foreach($key in 'Width','Height','Left','Top'){
            if([math]::Abs($a[$i].Note.$key-$b[$i].Note.$key) -gt 0.75){return $false}
            $a[$i].Note.PSObject.Properties.Remove($key);$b[$i].Note.PSObject.Properties.Remove($key)
        }}
    }
    return ($a|ConvertTo-Json -Depth 25 -Compress) -ceq ($b|ConvertTo-Json -Depth 25 -Compress)
}
$excel=$null;$book=$null
try{
    $originalHash=if($WorkbookPath){(Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$null}
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if($WorkbookPath){$book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath,$repo),0,$true)}
    else{
        $book=$excel.Workbooks.Add();$s=$book.Worksheets.Item(1);$s.Name='内容对照'
        try{
            Set-Text $s 'A1' '中文Red蓝色BoldItalic';Set-Text $s 'A2' '批注内容';Set-Text $s 'F1' '可编辑值'
            $cell=$s.Range('A1');$font=$cell.Font
            try{
                $font.Name='Calibri';$font.Size=11
                foreach($span in @(@(1,2,'Microsoft YaHei',14,255,$false,$false),@(3,3,'Arial',16,16711680,$true,$false),@(6,2,'Microsoft YaHei',12,32768,$false,$false),@(8,4,'Calibri',18,0,$true,$false),@(12,6,'Times New Roman',13,8388736,$false,$true))){
                    $chars=$cell.Characters($span[0],$span[1]);$f=$chars.Font
                    try{$f.Name=$span[2];$f.Size=$span[3];$f.Color=$span[4];$f.Bold=$span[5];$f.Italic=$span[6]}finally{Release-Com $f;Release-Com $chars}
                }
            }finally{Release-Com $font;Release-Com $cell}
            foreach($a in 'A1','A2'){$cell=$s.Range($a);try{$note=$cell.AddComment("审核：保留中文与换行`n第二行 $a");try{$note.Visible=($a -eq 'A2')}finally{Release-Com $note}}finally{Release-Com $cell}}
            foreach($entry in @(@('B1','https://example.invalid/path?q=1&x=2','','网页说明','示例链接'),@('C1','','''第二表''!$A$1','内部说明','转到第二表'),@('D1','mailto:sample@example.invalid','','邮件说明','邮件链接'))){
                $cell=$s.Range($entry[0]);$links=$s.Hyperlinks
                try{$link=$links.Add($cell,$entry[1],$entry[2],$entry[3],$entry[4]);Release-Com $link}finally{Release-Com $links;Release-Com $cell}
            }
            $column=$s.Range('A:A');$rows=$s.Range('1:2');try{$column.ColumnWidth=45;$rows.RowHeight=35}finally{Release-Com $column;Release-Com $rows}
        }finally{Release-Com $s}
        $s=$book.Worksheets.Add();try{$s.Name='第二表';Set-Text $s 'A1' '内部链接目标'}finally{Release-Com $s}
        $s=$book.Worksheets.Item('内容对照');try{$s.Activate()}finally{Release-Com $s}
    }
    # Locate by name: Excel inserts the added sheet before the original.
    $s=$book.Worksheets.Item('内容对照');$first=$book.Worksheets.Item(1)
    try{if($s.Index -ne 1){$s.Move($first)}}finally{Release-Com $first;Release-Com $s}
    $before=Capture $book;$calculation=$excel.Calculation
    $save=Join-Path $target 'excel-saved.xlsx';$book.SaveAs($save,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($save,0,$true);$after=Capture $book
    $same=Equivalent $before $after
    $matches=$null;if($ReferencePath){$reference=Get-Content -LiteralPath ([IO.Path]::GetFullPath($ReferencePath,$repo)) -Raw|ConvertFrom-Json;$matches=Equivalent $reference.$ReferenceState $after}
    $edited=$null
    if(-not $WorkbookPath){
        $s=$book.Worksheets.Item('内容对照');try{Set-Text $s 'B1' '修改后的链接文字';Set-Text $s 'A2' '修改批注所在格'}finally{Release-Com $s}
        $s=$book.Worksheets.Item('内容对照');$cell=$s.Range('A1');$font=$cell.Font
        try{$font.Bold=$true}finally{Release-Com $font;Release-Com $cell;Release-Com $s}
        $edited=Capture $book;$book.SaveAs((Join-Path $target 'excel-edited.xlsx'),51)
    }
    $unchanged=if($WorkbookPath){$originalHash -eq (Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$true}
    [ordered]@{ExcelVersion=$excel.Version;ExcelBuild=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=$calculation;GeometryTolerancePoints=0.75;Before=$before;AfterReopen=$after;AfterEdit=$edited;SameAfterReopen=$same;ReferenceMatches=$matches;SourceUnchanged=$unchanged}|ConvertTo-Json -Depth 25|Set-Content -LiteralPath (Join-Path $target 'excel-cell-text.json') -Encoding utf8
    if(-not $same -or $matches -eq $false -or -not $unchanged){throw 'Cell text reference comparison failed; evidence retained.'}
    Write-Output "PASS: Excel rich text, links, comments and reopen; $target"
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
