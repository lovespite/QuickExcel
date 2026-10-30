param([Parameter(Mandatory)][string]$OutputDirectory, [ValidateSet('ReadOnly','Unknown')][string]$Scenario='ReadOnly',
    [string]$WorkbookPath, [string]$ReferencePath, [ValidateSet('AfterReopen','AfterEdit')][string]$ReferenceState='AfterReopen')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($object){if($null -ne $object){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}}
function Set-Value($sheet,$address,$value){$cell=$sheet.Range($address);try{if($value -is [string]){$cell.Value2=[string]$value}else{$cell.Value2=[double]$value}}finally{Release-Com $cell}}
function Set-Formula($sheet,$address,$formula,$kind='Formula'){$cell=$sheet.Range($address);try{$cell.$kind=$formula}finally{Release-Com $cell}}
function Capture-Cells($book){
    $records=@()
    for($i=1;$i -le $book.Worksheets.Count;$i++){
        $sheet=$book.Worksheets.Item($i);$used=$sheet.UsedRange
        try{
            for($r=1;$r -le $used.Rows.Count;$r++){for($c=1;$c -le $used.Columns.Count;$c++){
                $cell=$used.Cells.Item($r,$c)
                try{
                    $arrayRange=$null
                    if($cell.HasArray){$array=$cell.CurrentArray;try{$arrayRange=[string]$array.Address($false,$false)}finally{Release-Com $array}}
                    $records += [ordered]@{Sheet=[string]$sheet.Name;Cell=[string]$cell.Address($false,$false);Value=$cell.Value2;Formula=$cell.Formula;Formula2=$cell.Formula2;Text=[string]$cell.Text;ArrayRange=$arrayRange}
                }finally{Release-Com $cell}
            }}
        }finally{Release-Com $used;Release-Com $sheet}
    }
    return ,$records
}
$excel=$null;$book=$null
try{
    $sourceHash=if($WorkbookPath){(Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$null}
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if($WorkbookPath){$book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath,$repo),0,$true)}
    else{
        $book=$excel.Workbooks.Add()
        while($book.Worksheets.Count -gt 1){$extra=$book.Worksheets.Item($book.Worksheets.Count);try{$extra.Delete()}finally{Release-Com $extra}}
        $sheet=$book.Worksheets.Item(1)
        try{
            if($Scenario -eq 'Unknown'){
                $sheet.Name='未知函数';Set-Value $sheet 'A1' ([double]9);Set-Formula $sheet 'B1' '=SQRT(A1)';Set-Formula $sheet 'C1' '=B1+1';Set-Formula $sheet 'D1' '=SUM(A1,2)'
            }else{
                $sheet.Name='传统数组';Set-Value $sheet 'A1' ([double]1);Set-Value $sheet 'A2' ([double]2);Set-Value $sheet 'A3' ([double]3)
                Set-Formula $sheet 'B1:B3' '=A1:A3*2' 'FormulaArray';Set-Formula $sheet 'D1' '=SUM(B1:B3)';Set-Formula $sheet 'E1' '=D1+1'
            }
        }finally{Release-Com $sheet}
        if($Scenario -eq 'ReadOnly'){
            foreach($name in @('结构化引用','动态数组','版本函数')){
                $last=$book.Worksheets.Item($book.Worksheets.Count);try{$sheet=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last}
                try{
                    $sheet.Name=$name
                    if($name -eq '结构化引用'){
                        Set-Value $sheet 'A1' 'Item';Set-Value $sheet 'B1' 'Amount'
                        for($i=2;$i -le 4;$i++){Set-Value $sheet ('A'+$i) ('item'+$i);Set-Value $sheet ('B'+$i) ([double](($i-1)*10))}
                        $range=$sheet.Range('A1:B4');$tables=$sheet.ListObjects
                        try{$table=$tables.Add(1,$range,[Type]::Missing,1);try{$table.Name='DataTable'}finally{Release-Com $table}}finally{Release-Com $tables;Release-Com $range}
                        Set-Formula $sheet 'D1' '=SUM(DataTable[Amount])';Set-Formula $sheet 'E1' '=D1+1'
                    }elseif($name -eq '动态数组'){Set-Formula $sheet 'B1' '=SEQUENCE(3,1,10,1)' 'Formula2';Set-Formula $sheet 'D1' '=SUM(B1:B3)'}
                    else{Set-Value $sheet 'A1' ([double]1);Set-Value $sheet 'A2' ([double]2);Set-Value $sheet 'B1' ([double]7);Set-Value $sheet 'B2' ([double]9);Set-Formula $sheet 'D1' '=XLOOKUP(2,A1:A2,B1:B2)' 'Formula2'}
                }finally{Release-Com $sheet}
            }
        }
    }
    $excel.CalculateFullRebuild();$calculation=[int]$excel.Calculation;$before=Capture-Cells $book
    $saved=Join-Path $target 'excel-source.xlsx';$book.SaveAs($saved,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($saved,0,$true);$excel.CalculateFullRebuild();$after=Capture-Cells $book
    $edited=$null
    if(-not $WorkbookPath -and $Scenario -eq 'Unknown'){
        $sheet=$book.Worksheets.Item(1);try{Set-Value $sheet 'A1' ([double]16)}finally{Release-Com $sheet}
        $excel.CalculateFullRebuild();$edited=Capture-Cells $book;$book.SaveAs((Join-Path $target 'excel-edited.xlsx'),51)
    }
    $book.Close($false);Release-Com $book;$book=$null
    $same=($before|ConvertTo-Json -Depth 8 -Compress) -ceq ($after|ConvertTo-Json -Depth 8 -Compress)
    $matches=if($ReferencePath){$reference=Get-Content -LiteralPath $ReferencePath -Raw|ConvertFrom-Json;($reference.$ReferenceState|ConvertTo-Json -Depth 8 -Compress) -ceq ($before|ConvertTo-Json -Depth 8 -Compress)}else{$null}
    $unchanged=if($WorkbookPath){$sourceHash -eq (Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$true}
    [ordered]@{ExcelVersion=[string]$excel.Version;ExcelBuild=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=$calculation;Scenario=$Scenario;
        Before=$before;AfterReopen=$after;AfterEdit=$edited;SameAfterReopen=$same;ReferenceMatches=$matches;SourceUnchanged=$unchanged;SourceHash=$sourceHash}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $target 'excel-formulas.json')
    if(-not $same -or -not $unchanged -or $matches -eq $false){throw 'Excel formula comparison failed; see JSON.'}
    Write-Output "PASS: Excel formulas, full calculation and reopen; $target"
}finally{if($book){$book.Close($false);Release-Com $book};if($excel){[GC]::Collect();[GC]::WaitForPendingFinalizers();$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
