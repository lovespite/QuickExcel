param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Put($s,[int]$r,[int]$c,$value){$cell=$s.Cells.Item($r,$c);try{
 if($null -eq $value){[void]$cell.ClearContents()}elseif($value -is [string]){$cell.NumberFormat='@';$cell.Value2=[string]$value}elseif($value -is [bool]){$cell.Value2=[bool]$value}else{$cell.Value2=[double]$value}
}finally{Release-Com $cell}}
function Formula($s,[int]$r,[int]$c,[string]$formula){$cell=$s.Cells.Item($r,$c);try{$cell.Formula=$formula}catch{throw "Formula rejected at $($s.Name) R${r}C${c}: $formula ($_)"}finally{Release-Com $cell}}
function New-Sheet([string]$name){$last=$book.Worksheets.Item($book.Worksheets.Count);try{$s=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last};$s.Name=$name;return $s}
function Capture($workbook){
 $result=@()
 for($i=1;$i -le $workbook.Worksheets.Count;$i++){$s=$workbook.Worksheets.Item($i);try{
  $cells=@();$rows=@();$cols=@();$used=$s.UsedRange;try{$rowCount=[Math]::Min(400,$used.Row+$used.Rows.Count);$colCount=[Math]::Min(14,$used.Column+$used.Columns.Count)}finally{Release-Com $used}
  for($r=1;$r -le $rowCount;$r++){for($c=1;$c -le $colCount;$c++){$cell=$s.Cells.Item($r,$c);try{
   $v=$cell.Value2;$f=if($cell.HasFormula){[string]$cell.Formula}else{$null};$kind=if($null -eq $v){'Blank'}elseif($v -is [string]){'Text'}elseif($v -is [bool]){'Boolean'}else{'Number'}
   if($f -and [string]$cell.Text -in @('#REF!','#VALUE!','#DIV/0!','#N/A','#NAME?','#NUM!','#NULL!')){$kind='Error';$v=[string]$cell.Text}
   if($null -ne $v -or $f){$cells+=@{Address=$cell.Address($false,$false);Kind=$kind;Value=$v;Formula=$f;Text=[string]$cell.Text}}
  }finally{Release-Com $cell}}}
  $outline=$s.Outline;try{$below=$outline.SummaryRow -eq 1;$right=$outline.SummaryColumn -eq -4152}finally{Release-Com $outline}
  foreach($r in 1..40){$row=$s.Rows.Item($r);$adjacent=$null;try{
   $neighbor=if($below){$r-1}else{$r+1};$collapsed=$false
   if($neighbor -ge 1){$adjacent=$s.Rows.Item($neighbor);if($adjacent.OutlineLevel -gt $row.OutlineLevel){$collapsed=-not [bool]$row.ShowDetail}}
   $rows+=@{Index=$r;Level=$row.OutlineLevel;Hidden=[bool]$row.Hidden;Collapsed=$collapsed}
  }finally{Release-Com $adjacent;Release-Com $row}}
  foreach($c in 1..14){$col=$s.Columns.Item($c);$adjacent=$null;try{
   $neighbor=if($right){$c-1}else{$c+1};$collapsed=$false
   if($neighbor -ge 1){$adjacent=$s.Columns.Item($neighbor);if($adjacent.OutlineLevel -gt $col.OutlineLevel){$collapsed=-not [bool]$col.ShowDetail}}
   $cols+=@{Index=$c;Level=$col.OutlineLevel;Hidden=[bool]$col.Hidden;Collapsed=$collapsed}
  }finally{Release-Com $adjacent;Release-Com $col}}
  $outline=$s.Outline;try{$result+=@{Sheet=$s.Name;Cells=$cells;Rows=$rows;Columns=$cols;SummaryBelow=($outline.SummaryRow -eq 1);SummaryRight=($outline.SummaryColumn -eq -4152)}}finally{Release-Com $outline}
 }finally{Release-Com $s}};return ,$result
}
$excel=$null;$book=$null;$stages=@();$cases=@();$rejected=@()
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('OutlineExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OutlineExcelOwner {[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);}'}
 [uint32]$ownedExcelProcess=0;[void][OutlineExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
 if($WorkbookPath){$source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true);$names=@('output')}
 else{
  $book=$excel.Workbooks.Add(-4167);$s=$book.Worksheets.Item(1);$s.Name='Subtotals'
  try{
   Put $s 1 1 'Value';Put $s 1 2 'Keep';$values=@(10,'20',$true,$null,'',-3,0,2,8,5,4)
   for($i=0;$i -lt $values.Count;$i++){Put $s ($i+2) 1 $values[$i];Put $s ($i+2) 2 $(if($i -in @(0,6,8)){'no'}else{'yes'})}
   Formula $s 6 1 '=""';Formula $s 7 1 '=1/0';Formula $s 10 1 '=SUBTOTAL(9,A2:A5)';Formula $s 11 1 '=1+SUBTOTAL(9,A2:A5)';Formula $s 12 1 '=SUM(A2,A9)'
   Put $s 2 3 2;Put $s 2 4 3;Put $s 2 5 4
   $row=1
   foreach($fn in (1..11 + 101..111)){foreach($range in @('A2:A5','A2:A9','A2:A12','A13:A20','A9:A9','C2:E2','A2,A9,A2','A:A')){Formula $s $row 7 ('=SUBTOTAL('+$fn+','+$range+')');$row++}}
   foreach($formula in @('=SUBTOTAL(9.9,A2:A5)','=SUBTOTAL("9",A2:A5)','=SUBTOTAL(TRUE,A2:A5)','=SUBTOTAL(0,A2:A5)','=SUBTOTAL(12,A2:A5)','=SUBTOTAL(100,A2:A5)','=SUBTOTAL(112,A2:A5)','=SUBTOTAL(-1,A2:A5)','=SUBTOTAL(9,10)','=SUBTOTAL(9,"20")','=SUBTOTAL(9,TRUE)','=SUBTOTAL(9,A2+1)','=SUBTOTAL(9,IF(TRUE,A2,A9))','=SUBTOTAL(9,INDEX(A2:A5,1))','=SUBTOTAL(9,IF(FALSE,A2,A9))','=SUBTOTAL(9,#N/A)','=SUBTOTAL(9,)','=SUBTOTAL(,A2:A5)')){try{Formula $s $row 7 $formula}catch{$rejected+=@{Formula=$formula;Message=[string]$_}};$row++}
  }finally{Release-Com $s}
  foreach($name in @('OutlineBelow','OutlineAbove')){$s=New-Sheet $name;try{
   $outline=$s.Outline;try{$outline.SummaryRow=if($name -eq 'OutlineAbove'){0}else{1};$outline.SummaryColumn=if($name -eq 'OutlineAbove'){-4131}else{-4152}}finally{Release-Com $outline}
   foreach($r in 1..16){Put $s $r 1 ('row'+$r);Put $s $r 2 $r}
   Formula $s 20 2 '=SUBTOTAL(9,B2:B16)';Formula $s 21 2 '=SUBTOTAL(109,B2:B16)'
   foreach($range in @('3:5','8:10','2:12')){$rg=$s.Rows.Item($range);try{[void]$rg.Group()}finally{Release-Com $rg}}
   foreach($range in @('C:E','B:G')){$rg=$s.Columns.Item($range);try{[void]$rg.Group()}finally{Release-Com $rg}}
  }finally{Release-Com $s}}
  $codes=@(-4106,-4113,-4112,-4136,-4139,-4149,-4155,-4156,-4157,-4164,-4165)
  for($fn=1;$fn -le 11;$fn++){$s=New-Sheet ('Generate'+$fn);try{
   Put $s 1 1 'Group';Put $s 1 2 'Amount';Put $s 1 3 'Other'
   $groups=@('a','A','b','b','b','c','a')
   for($i=0;$i -lt $groups.Count;$i++){Put $s ($i+2) 1 $groups[$i];Put $s ($i+2) 2 ($i+1);Formula $s ($i+2) 3 ('=B'+($i+2)+'*2')}
   Formula $s 1 5 '=B8';Put $s 10 1 'tail';$cases+=@{Sheet=$s.Name;Range='A1:C8';Function=$fn;ExcelFunction=$codes[$fn-1];Group=1;Totals=@(2,3)}
  }finally{Release-Com $s}}
  for($si=1;$si -le $book.Worksheets.Count;$si++){$s=$book.Worksheets.Item($si);$columns=$s.Range('A:N');try{$columns.ColumnWidth=24}finally{Release-Com $columns;Release-Com $s}}
  $source=Join-Path $target 'source.xlsx';$excel.CalculateFullRebuild();$book.SaveAs($source,51);$book.Close($false);Release-Com $book;$book=$null
  $sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$false);$names=@('initial','manual','filtered','collapsed','expanded','generated')
 }
 foreach($name in $names){
  if($name -eq 'manual'){$s=$book.Worksheets.Item('Subtotals');$r=$s.Rows.Item(2);$c=$s.Columns.Item('D');try{$r.Hidden=$true;$c.Hidden=$true}finally{Release-Com $r;Release-Com $c;Release-Com $s}}
  if($name -eq 'filtered'){$s=$book.Worksheets.Item('Subtotals');$r=$s.Range('A1:B12');try{[void]$r.AutoFilter(2,'yes')}finally{Release-Com $r;Release-Com $s}}
  if($name -eq 'collapsed'){
   foreach($sheetName in @('OutlineBelow','OutlineAbove')){$s=$book.Worksheets.Item($sheetName);try{
    foreach($r in $(if($sheetName -eq 'OutlineBelow'){@(6,13)}else{@(2,1)})){$row=$s.Rows.Item($r);try{$row.ShowDetail=$false}finally{Release-Com $row}}
    $col=$s.Columns.Item($(if($sheetName -eq 'OutlineBelow'){'H'}else{'A'}));try{$col.ShowDetail=$false}finally{Release-Com $col}
   }finally{Release-Com $s}}
  }
  if($name -eq 'expanded'){
   foreach($sheetName in @('OutlineBelow','OutlineAbove')){$s=$book.Worksheets.Item($sheetName);try{$row=$s.Rows.Item($(if($sheetName -eq 'OutlineBelow'){13}else{1}));$col=$s.Columns.Item($(if($sheetName -eq 'OutlineBelow'){'H'}else{'A'}));try{$row.ShowDetail=$true;$col.ShowDetail=$true}finally{Release-Com $row;Release-Com $col}}finally{Release-Com $s}}
  }
  if($name -eq 'generated'){foreach($spec in $cases){$s=$book.Worksheets.Item($spec.Sheet);$r=$s.Range($spec.Range);try{[void]$r.Subtotal($spec.Group,$spec.ExcelFunction,[object[]]$spec.Totals,$true,$false,1)}finally{Release-Com $r;Release-Com $s}}}
  $excel.CalculateFullRebuild();$before=Capture $book;$path=Join-Path $target ($name+'.xlsx');$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
  $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($path,0,$false);$excel.CalculateFullRebuild();$after=Capture $book
  $stable=(($before|ConvertTo-Json -Depth 12 -Compress)-eq($after|ConvertTo-Json -Depth 12 -Compress));if(-not $stable){throw ('Unstable stage '+$name)}
  $stages+=@{Name=$name;Path=$path;Sha256=$hash;Sheets=$after;Stable=$stable}
 }
 @{Source=$source;SourceSha256=$sourceHash;SourceUnchanged=($sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash);Cases=$cases;Rejected=$rejected;Stages=$stages}|ConvertTo-Json -Depth 15|Set-Content (Join-Path $target 'outline-subtotal.json') -Encoding utf8
}catch{
 ($_ | Out-String)+$_.ScriptStackTrace | Set-Content (Join-Path $target 'error.txt') -Encoding utf8;throw
}finally{
 if($book){$book.Close($false);Release-Com $book;$book=$null};[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
