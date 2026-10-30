param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath,[switch]$Date1904)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Set-Value($cell,$value){
 if($null -eq $value){[void]$cell.ClearContents()}
 elseif($value -is [string]){$cell.NumberFormat='@';$cell.Value2=[string]$value}
 elseif($value -is [bool]){$cell.Value2=[bool]$value}
 else{$cell.Value2=[double]$value}
}
function Sheet([string]$name){
 if($book.Worksheets.Count -eq 1 -and $book.Worksheets.Item(1).Name -eq 'Sheet1'){$s=$book.Worksheets.Item(1)}
 else{$last=$book.Worksheets.Item($book.Worksheets.Count);try{$s=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last}}
 $s.Name=$name;return $s
}
function Put($s,[int]$row,[int]$column,$value){$c=$s.Cells.Item($row,$column);try{Set-Value $c $value}finally{Release-Com $c}}
function Formula($s,[int]$row,[int]$column,[string]$formula){$c=$s.Cells.Item($row,$column);try{$c.Formula=$formula}finally{Release-Com $c}}
function Capture($workbook){
 $result=@()
 for($si=1;$si -le $workbook.Worksheets.Count;$si++){
  $s=$workbook.Worksheets.Item($si);try{
   $cells=@()
   for($r=1;$r -le 32;$r++){for($c=1;$c -le 8;$c++){
    $cell=$s.Cells.Item($r,$c);try{
     $v=$cell.Value2;$formula=if($cell.HasFormula){[string]$cell.Formula}else{$null}
     $kind=if($null -eq $v){'Blank'}elseif($v -is [string]){'Text'}elseif($v -is [bool]){'Boolean'}else{'Number'}
     if($formula -and [string]$cell.Text -in @('#REF!','#VALUE!','#DIV/0!','#N/A','#NAME?','#NUM!','#NULL!')){$kind='Error';$v=[string]$cell.Text}
     $cells+=@{Address=$cell.Address($false,$false);Kind=$kind;Value=$v;Formula=$formula;Text=[string]$cell.Text;NumberFormat=$cell.NumberFormat}
    }finally{Release-Com $cell}
   }}
   $result+=@{Sheet=$s.Name;Cells=$cells}
  }finally{Release-Com $s}
 };return ,$result
}
$excel=$null;$book=$null;$cases=@()
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('CleaningExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class CleaningExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
 [uint32]$ownedExcelProcess=0;[void][CleaningExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
 if($WorkbookPath){$source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true)}
 else{
  $book=$excel.Workbooks.Add(-4167);$book.Date1904=[bool]$Date1904
  $s=Sheet 'Duplicates'
  try{
   foreach($col in 1..4){Put $s 1 $col ('column'+$col)}
   $values=@('red','RED',1,'1','001','01',1,1,$null,'',$true,'TRUE','*','*','?','?',0,'0',0.01,'1%',3,'3.0',3,'3e0')
   for($i=0;$i -lt $values.Count;$i++){ $r=$i+2;Put $s $r 1 $values[$i];Put $s $r 2 ('record'+$r);Formula $s $r 3 ('=A'+$r);Put $s $r 4 ('outside'+$r) }
   foreach($pair in @(@(8,'000'),@(9,'0.00'))){$c=$s.Cells.Item($pair[0],1);try{$c.NumberFormat=$pair[1]}finally{Release-Com $c}}
   Formula $s 1 6 '=A4';Formula $s 2 6 '=A8';Formula $s 3 6 '=A25';Formula $s 4 6 '=SUM(A2:A25)'
   Formula $s 26 1 '=1/0';Put $s 26 2 'error-number';Formula $s 27 1 '=1/0';Put $s 27 2 'error-duplicate';Put $s 28 1 '#DIV/0!';Put $s 28 2 'error-text'
   $cases+=@{Sheet='Duplicates';Operation='Duplicates';Range='A1:C28';Keys=@(1);Header=$true}
  }finally{Release-Com $s}
  $s=Sheet 'MultipleKeys'
  try{
   Put $s 1 1 'Name';Put $s 1 2 'Group';Put $s 1 3 'Record'
   $rows=@(@('a',1),@('A',1),@('a',2),@('b',2),@('B','2'),@('b','02'),@('c',$null),@('c',''),@('c',$null))
   for($i=0;$i -lt $rows.Count;$i++){Put $s ($i+2) 1 $rows[$i][0];Put $s ($i+2) 2 $rows[$i][1];Put $s ($i+2) 3 ($i+1)}
   $cases+=@{Sheet='MultipleKeys';Operation='Duplicates';Range='A1:C10';Keys=@(1,2);Header=$true}
  }finally{Release-Com $s}
  $s=Sheet 'FormulaKeys'
  try{
   foreach($i in 1..6){Formula $s $i 1 $(if($i -lt 4){'=1+0'}else{'=2+0'});Formula $s $i 2 ('=A'+$i+'*10');Put $s $i 3 $i}
   Formula $s 1 5 '=A4';Formula $s 2 5 '=B6'
   $cases+=@{Sheet='FormulaKeys';Operation='Duplicates';Range='A1:C6';Keys=@(1);Header=$false}
  }finally{Release-Com $s}
  foreach($spec in @(
   @{Name='DelimitedText';Values=@('001,2.50,2026-10-31','"A,B",007,"2026-10-01"','a,,c',',b,','"say ""Hi""",02,x',"'001,003,b",'a,b','',$null,'=1+1,004,z','中,😀,0');Delimiter=',';Formats=@(2,2,2);Collapse=$false},
   @{Name='Collapsed';Values=@('a,,,b,',' a , b ','a,,c','"a,b",c','a,"b""c",d');Delimiter=',';Formats=@(2,2,2);Collapse=$true},
   @{Name='NumbersAndDates';Values=@('001;1,234.50;2026-10-31','002;-12.5;2026-01-02','010;12-;1900-02-28','020;50%;2026-02-28','030;1e3;2026-12-31');Delimiter=';';Formats=@(2,1,5);Collapse=$false},
   @{Name='NumericChoice';Values=@('001','123','1,234.50','-12.5','12-','1e3','0.5','50%','123456789012345','1234567890123456','2026-10-31','03/04/2026','TRUE','1/2');Delimiter=';';Formats=@(1);Collapse=$false},
   @{Name='DateDMY';Values=@('31/10/2026','02/01/2026','29/02/1900','01/03/1900');Delimiter=';';Formats=@(4);Collapse=$false},
   @{Name='FixedText';Values=@('001ABCD2026-10-31','002EFGH2026-01-02','abcdef','ab','中甲乙中文汉字尾','a😀bcdefg');Fixed=@(0,3,7);Formats=@(2,2,2)}
  )){
   $s=Sheet $spec.Name
   try{
    Put $s 1 1 'Source'
    for($i=0;$i -lt $spec.Values.Count;$i++){
     $inputValue=$spec.Values[$i]
     if($Date1904 -and $inputValue -is [string] -and $spec.Name -in @('NumbersAndDates','DateDMY')){$inputValue=$inputValue.Replace('1900','1904')}
     Put $s ($i+2) 1 $inputValue
    }
    $range='A2:A'+($spec.Values.Count+1)
    $cases+=@{Sheet=$spec.Name;Operation=if($spec.Fixed){'Fixed'}else{'Delimited'};Range=$range;Delimiter=$spec.Delimiter;Collapse=[bool]$spec.Collapse;Formats=$spec.Formats;Fixed=$spec.Fixed;Destination='C2';ExplicitLongIntegerDisplay=($spec.Name -eq 'NumericChoice')}
   }finally{Release-Com $s}
  }
  for($si=1;$si -le $book.Worksheets.Count;$si++){$s=$book.Worksheets.Item($si);$columns=$s.Columns;try{$columns.ColumnWidth=24}finally{Release-Com $columns;Release-Com $s}}
  $excel.CalculateFullRebuild();$before=Capture $book;$source=Join-Path $target 'source.xlsx';$book.SaveAs($source,51)
  $book.Close($false);Release-Com $book;$book=$null;$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
  $book=$excel.Workbooks.Open($source,0,$false)
  foreach($spec in $cases){
   $s=$book.Worksheets.Item($spec.Sheet);$range=$s.Range($spec.Range);$destination=$null
   try{
    if($spec.Operation -eq 'Duplicates'){$keys=[object[]]$spec.Keys;[void]$range.RemoveDuplicates($keys,$(if($spec.Header){1}else{2}))}
    else{
     $destination=$s.Range($spec.Destination);$fields=New-Object 'System.Collections.Generic.List[object]'
     for($i=0;$i -lt $spec.Formats.Count;$i++){$fields.Add([object[]]@($(if($spec.Fixed){$spec.Fixed[$i]}else{$i+1}),$spec.Formats[$i]))}
     [void]$range.TextToColumns($destination,$(if($spec.Fixed){2}else{1}),1,[bool]$spec.Collapse,$false,($spec.Delimiter -eq ';'),($spec.Delimiter -eq ','),$false,$false,[Type]::Missing,$fields.ToArray(),'.',',',$true)
     # This is an explicit Number choice, not Excel's automatic General inference:
     # keep the 15-digit integer fully visible in the same way as the QuickExcel preview.
     if($spec.ExplicitLongIntegerDisplay){$long=$s.Range('C10');try{$long.NumberFormat='0'}finally{Release-Com $long}}
    }
   }finally{Release-Com $destination;Release-Com $range;Release-Com $s}
  }
 }
 $excel.CalculateFullRebuild();$after=Capture $book;$output=Join-Path $target 'output.xlsx';$book.SaveAs($output,51);$book.Close($false);Release-Com $book;$book=$null
 $hash=(Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($output,0,$false);$excel.CalculateFullRebuild();$reopened=Capture $book
 $stable=(($after|ConvertTo-Json -Depth 12 -Compress)-eq($reopened|ConvertTo-Json -Depth 12 -Compress))
 @{Source=$source;SourceSha256=$sourceHash;SourceUnchanged=$sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;Path=$output;Sha256=$hash;Date1904=[bool]$book.Date1904;
   Cases=$cases;Before=$before;After=$after;Reopened=$reopened;Stable=$stable}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $target 'cleaning.json') -Encoding utf8
 if(-not $stable){throw 'Data cleaning changed after save/reopen.'}
}finally{
 if($book){$book.Close($false);Release-Com $book;$book=$null};[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
